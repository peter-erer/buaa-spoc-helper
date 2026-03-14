// ==UserScript==
// @name         智学北航学习助手
// @namespace    https://github.com/peter-erer/buaa-spoc-helper
// @version      5.1.0
// @description  支持 MSA 回放字幕/PPT 导出、课程详情视频/PPTX 下载，以及 SPOC 播放器辅助工具
// @author       micraow(github.com/micraow),Peter-erer(github.com/peter-erer)
// @license      MIT
// @match        *://*.classroom.msa.buaa.edu.cn/livingroom*
// @match        *://*.classroom.msa.buaa.edu.cn/coursedetail*
// @match        *://spoc.buaa.edu.cn/*
// @homepageURL  https://github.com/peter-erer/buaa-spoc-helper
// @supportURL   https://github.com/peter-erer/buaa-spoc-helper/issues
// @updateURL    https://github.com/peter-erer/buaa-spoc-helper/raw/main/buaa-spoc-helper.js
// @downloadURL  https://github.com/peter-erer/buaa-spoc-helper/raw/main/buaa-spoc-helper.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=buaa.edu.cn
// @require      https://cdn.jsdelivr.net/npm/pptxgenjs@3.9.0/dist/pptxgen.bundle.js
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  if (window.__BUAA_SPOC_HELPER_LOADED__) return;
  window.__BUAA_SPOC_HELPER_LOADED__ = true;

  const SPOC_RATE_PRESETS = [1, 1.25, 1.5, 2];
  const COURSE_READY_STATUS = "6";

  const state = {
    mode: detectMode(),
    livingroom: {
      subtitleData: null,
      pptData: null,
      ui: null,
      interceptorInstalled: false,
    },
    courseDetail: {
      courseId: getQueryParam("course_id"),
      token: "",
      account: "",
      courseData: null,
      userInfoPromise: null,
      ui: null,
      sessions: [],
      rowCount: 0,
    },
    spoc: {
      ui: null,
      logs: [],
      autoNextEnabled: false,
      preferredRate: 1.5,
      interceptorInstalled: false,
      isRunning: false,
      fastMode: false,
      simSpeed: 10,
      simInterval: 1,
      simDuration: 0,
      simProgress: 0,
      simPlayTime: 0,
      simTickBusy: false,
      simStatusText: "",
      simInfoText: "",
      timer: null,
      params: null,
      capturedHeaders: {},
      currentVideo: null,
      currentDuration: 0,
      currentSrc: "",
      refreshTimer: null,
      rateEnforceTimer: null,
      currentResourceTitle: "",
    },
  };

  ready(() => {
    injectSharedStyles();

    switch (state.mode) {
      case "livingroom":
        initLivingroom();
        break;
      case "coursedetail":
        initCourseDetail();
        break;
      case "spoc":
        initSpocHelper();
        break;
      default:
        break;
    }
  });

  function detectMode() {
    const { hostname, pathname } = window.location;
    if (hostname.includes("classroom.msa.buaa.edu.cn")) {
      if (pathname.includes("/livingroom")) return "livingroom";
      if (pathname.includes("/coursedetail")) return "coursedetail";
    }
    if (hostname === "spoc.buaa.edu.cn") {
      if (pathname.includes("/spocnew/mycourse/coursecenter/")) return "spoc";
      return "unsupported";
    }
    return "unsupported";
  }

  function ready(callback) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", callback, { once: true });
    } else {
      callback();
    }
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function getQueryParam(key) {
    return new URLSearchParams(window.location.search).get(key) || "";
  }

  function getCourseCenterUrl() {
    return "https://classroom.msa.buaa.edu.cn/courseCenter";
  }

  function getCurrentCourseDetailUrl() {
    const courseId = getQueryParam("course_id");
    if (!courseId) return "";
    return `https://classroom.msa.buaa.edu.cn/coursedetail?course_id=${encodeURIComponent(courseId)}`;
  }

  function escapeHtml(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function sanitizeFilename(text, max = 90) {
    return String(text || "未命名")
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, max);
  }

  function normalizeText(text) {
    return String(text || "")
      .replace(/[\s\n\r\t]+/g, "")
      .replace(/[：:·•\-—_|（）()\[\]【】]/g, "")
      .toLowerCase();
  }

  function formatSrtTime(seconds) {
    const safeSeconds = Number.isFinite(seconds) ? Math.max(0, seconds) : 0;
    const totalMilliseconds = Math.floor(safeSeconds * 1000);
    const hours = Math.floor(totalMilliseconds / 3600000)
      .toString()
      .padStart(2, "0");
    const minutes = Math.floor((totalMilliseconds % 3600000) / 60000)
      .toString()
      .padStart(2, "0");
    const secs = Math.floor((totalMilliseconds % 60000) / 1000)
      .toString()
      .padStart(2, "0");
    const milliseconds = Math.floor(totalMilliseconds % 1000)
      .toString()
      .padStart(3, "0");
    return `${hours}:${minutes}:${secs},${milliseconds}`;
  }

  function formatClockTime(seconds) {
    const safeSeconds = Number.isFinite(seconds) ? Math.max(0, Math.floor(seconds)) : 0;
    const hours = Math.floor(safeSeconds / 3600)
      .toString()
      .padStart(2, "0");
    const minutes = Math.floor((safeSeconds % 3600) / 60)
      .toString()
      .padStart(2, "0");
    const secs = Math.floor(safeSeconds % 60)
      .toString()
      .padStart(2, "0");
    return `${hours}:${minutes}:${secs}`;
  }

  function formatDurationLabel(seconds) {
    const safe = Number.isFinite(seconds) ? Math.max(0, Math.floor(seconds)) : 0;
    const minutes = Math.floor(safe / 60);
    const secs = safe % 60;
    return `${minutes}分 ${secs.toString().padStart(2, "0")}秒`;
  }

  function downloadTextFile(content, fileName) {
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function triggerUrlDownload(url, fileName) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.rel = "noopener noreferrer";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  function openHtmlInPopup(html, target = "_blank") {
    const popup = window.open("", target);
    if (!popup) return null;

    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const blobUrl = URL.createObjectURL(blob);
    popup.location.replace(blobUrl);
    setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
    return popup;
  }

  function setBusy(button, busy, busyText) {
    if (!button) return;
    if (busy) {
      if (!button.dataset.originalText) {
        button.dataset.originalText = button.textContent || "";
      }
      button.disabled = true;
      if (busyText) button.textContent = busyText;
      button.dataset.busy = "true";
    } else {
      button.disabled = false;
      if (button.dataset.originalText) {
        button.textContent = button.dataset.originalText;
      }
      delete button.dataset.busy;
    }
  }

  async function waitFor(getter, { timeout = 20000, interval = 300 } = {}) {
    const startedAt = Date.now();
    while (Date.now() - startedAt <= timeout) {
      const value = getter();
      if (value) return value;
      await sleep(interval);
    }
    return null;
  }

  function makeDraggable(element, handle = element) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let initialLeft = 0;
    let initialTop = 0;

    const onMouseMove = (event) => {
      if (!dragging) return;
      const maxLeft = Math.max(8, window.innerWidth - element.offsetWidth - 8);
      const maxTop = Math.max(8, window.innerHeight - element.offsetHeight - 8);
      const nextLeft = Math.min(
        maxLeft,
        Math.max(8, initialLeft + (event.clientX - startX))
      );
      const nextTop = Math.min(
        maxTop,
        Math.max(8, initialTop + (event.clientY - startY))
      );
      element.style.left = `${nextLeft}px`;
      element.style.top = `${nextTop}px`;
      element.style.right = "auto";
    };

    const onMouseUp = () => {
      dragging = false;
      element.classList.remove("bsh-panel--dragging");
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    };

    handle.addEventListener("mousedown", (event) => {
      const target = event.target;
      if (
        target instanceof HTMLElement &&
        target.closest("button, input, textarea, a, select, label")
      ) {
        return;
      }
      dragging = true;
      startX = event.clientX;
      startY = event.clientY;
      initialLeft = element.offsetLeft;
      initialTop = element.offsetTop;
      element.classList.add("bsh-panel--dragging");
      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });
  }

  function getCookieValue(name) {
    const decoded = decodeURIComponent(document.cookie || "");
    const match = decoded.match(new RegExp(`(?:^|;\\s*)"?${name}"?=([^;]+)`));
    return match ? match[1].replace(/^"|"$/g, "") : "";
  }

  function getJwtTokenFromCookieToken() {
    const decoded = decodeURIComponent(document.cookie || "");
    const jwtMatch = decoded.match(/eyJ[\w-]+\.[\w-]+\.[\w-]+/);
    return jwtMatch ? jwtMatch[0] : "";
  }

  function getJwtAccountFromCookieToken() {
    const jwt = getJwtTokenFromCookieToken();
    if (jwt) {
      const parts = jwt.split(".");
      if (parts.length >= 2) {
        try {
          const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
          const padded = base64.padEnd(base64.length + ((4 - (base64.length % 4)) % 4), "=");
          const payload = JSON.parse(atob(padded));
          if (payload.account) {
            return String(payload.account);
          }
        } catch (error) {
          // ignore and fall through to JWTUser cookie
        }
      }
    }

    const jwtUser = getCookieValue("JWTUser");
    if (!jwtUser) return "";
    try {
      const payload = JSON.parse(jwtUser);
      return String(payload.account || "");
    } catch (error) {
      return "";
    }
  }

  function createPanelShell(id, title, subtitle) {
    const existing = document.getElementById(id);
    if (existing) existing.remove();

    const panel = document.createElement("section");
    panel.id = id;
    panel.className = "bsh-panel";
    panel.innerHTML = `
      <div class="bsh-panel__header">
        <div>
          <div class="bsh-panel__title">${escapeHtml(title)}</div>
          <div class="bsh-panel__subtitle">${escapeHtml(subtitle)}</div>
        </div>
        <span class="bsh-panel__drag">拖动</span>
      </div>
      <div class="bsh-panel__body"></div>
    `;
    document.body.appendChild(panel);
    makeDraggable(panel, panel.querySelector(".bsh-panel__header"));
    return {
      panel,
      body: panel.querySelector(".bsh-panel__body"),
    };
  }

  function injectSharedStyles() {
    if (document.getElementById("bsh-shared-styles")) return;

    const style = document.createElement("style");
    style.id = "bsh-shared-styles";
    style.textContent = `
      :root {
        --bsh-bg: rgba(12, 18, 32, 0.88);
        --bsh-bg-strong: rgba(8, 13, 24, 0.94);
        --bsh-card: rgba(255, 255, 255, 0.06);
        --bsh-card-strong: rgba(255, 255, 255, 0.1);
        --bsh-border: rgba(124, 168, 255, 0.18);
        --bsh-border-strong: rgba(124, 168, 255, 0.34);
        --bsh-primary: #4f8cff;
        --bsh-primary-strong: #2e74ff;
        --bsh-text: #f3f7ff;
        --bsh-text-muted: rgba(243, 247, 255, 0.72);
        --bsh-success: #41c88a;
        --bsh-warning: #f0b24c;
        --bsh-danger: #ff6b6b;
        --bsh-shadow: 0 16px 40px rgba(0, 0, 0, 0.24);
        --bsh-radius: 16px;
      }

      .bsh-panel {
        position: fixed;
        top: 120px;
        right: 24px;
        z-index: 99999;
        width: 336px;
        color: var(--bsh-text);
        background:
          linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.01)),
          linear-gradient(135deg, rgba(79, 140, 255, 0.12), rgba(8, 13, 24, 0.2) 40%),
          var(--bsh-bg);
        border: 1px solid var(--bsh-border);
        border-radius: var(--bsh-radius);
        box-shadow: var(--bsh-shadow);
        backdrop-filter: blur(16px);
        overflow: hidden;
        font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      }

      .bsh-panel--dragging {
        box-shadow: 0 22px 52px rgba(0, 0, 0, 0.28);
      }

      .bsh-panel__header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        padding: 16px 18px 14px;
        cursor: move;
        background: linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0));
        border-bottom: 1px solid rgba(255,255,255,0.06);
      }

      .bsh-panel__title {
        font-size: 15px;
        font-weight: 700;
        letter-spacing: 0.3px;
      }

      .bsh-panel__subtitle {
        margin-top: 4px;
        font-size: 12px;
        color: var(--bsh-text-muted);
      }

      .bsh-panel__drag {
        flex: none;
        font-size: 11px;
        color: var(--bsh-text-muted);
        padding: 5px 10px;
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 999px;
      }

      .bsh-panel__body {
        padding: 16px 18px 18px;
      }

      .bsh-section {
        margin-top: 14px;
      }

      .bsh-section:first-child {
        margin-top: 0;
      }

      .bsh-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
      }

      .bsh-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 10px;
        border-radius: 999px;
        border: 1px solid rgba(255,255,255,0.08);
        background: rgba(255,255,255,0.05);
        color: var(--bsh-text-muted);
        font-size: 11px;
        line-height: 1;
      }

      .bsh-badge::before {
        content: "";
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: rgba(255,255,255,0.35);
      }

      .bsh-badge.is-ready {
        color: var(--bsh-text);
        background: rgba(65, 200, 138, 0.12);
        border-color: rgba(65, 200, 138, 0.28);
      }

      .bsh-badge.is-ready::before {
        background: var(--bsh-success);
      }

      .bsh-badge.is-warning {
        background: rgba(240, 178, 76, 0.12);
        border-color: rgba(240, 178, 76, 0.28);
        color: var(--bsh-text);
      }

      .bsh-badge.is-warning::before {
        background: var(--bsh-warning);
      }

      .bsh-card {
        padding: 12px 14px;
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.08);
        background: var(--bsh-card);
      }

      .bsh-card--status {
        font-size: 12px;
        line-height: 1.6;
        color: var(--bsh-text-muted);
      }

      .bsh-card--grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 10px;
      }

      .bsh-metric {
        padding: 12px;
        border-radius: 12px;
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.06);
      }

      .bsh-metric__label {
        font-size: 11px;
        color: var(--bsh-text-muted);
      }

      .bsh-metric__value {
        margin-top: 6px;
        font-size: 14px;
        font-weight: 700;
        color: var(--bsh-text);
        word-break: break-all;
      }

      .bsh-actions,
      .bsh-rate-presets,
      .bsh-toolbar__actions,
      .bsh-inline-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
      }

      .bsh-btn {
        appearance: none;
        border: 1px solid transparent;
        border-radius: 10px;
        padding: 9px 12px;
        font-size: 12px;
        line-height: 1;
        cursor: pointer;
        transition: 0.18s ease;
        color: var(--bsh-text);
        background: rgba(255,255,255,0.06);
      }

      .bsh-btn:hover:not(:disabled) {
        transform: translateY(-1px);
        border-color: rgba(255,255,255,0.14);
        background: rgba(255,255,255,0.1);
      }

      .bsh-btn:disabled {
        cursor: not-allowed;
        opacity: 0.45;
      }

      .bsh-btn--primary {
        color: white;
        background: linear-gradient(135deg, var(--bsh-primary), var(--bsh-primary-strong));
        border-color: rgba(255,255,255,0.12);
      }

      .bsh-btn--ghost {
        background: rgba(255,255,255,0.03);
        border-color: rgba(255,255,255,0.08);
      }

      .bsh-btn--warn {
        background: rgba(240, 178, 76, 0.12);
        border-color: rgba(240, 178, 76, 0.25);
      }

      .bsh-btn--active {
        border-color: rgba(79, 140, 255, 0.38);
        background: rgba(79, 140, 255, 0.18);
      }

      .bsh-hints {
        margin: 0;
        padding-left: 18px;
        color: var(--bsh-text-muted);
        font-size: 12px;
        line-height: 1.6;
      }

      .bsh-hints li + li {
        margin-top: 6px;
      }

      .bsh-hints a {
        color: #b7d4ff;
        text-decoration: underline;
        text-underline-offset: 2px;
      }

      .bsh-hints a:hover {
        color: #ffffff;
      }

      .bsh-progress {
        overflow: hidden;
        height: 14px;
        border-radius: 999px;
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.08);
      }

      .bsh-progress__fill {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        min-width: 40px;
        height: 100%;
        padding-right: 8px;
        color: #fff;
        font-size: 10px;
        background: linear-gradient(90deg, var(--bsh-primary), #65b2ff);
        transition: width 0.24s ease;
      }

      .bsh-field-row {
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .bsh-input {
        width: 90px;
        padding: 8px 10px;
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,0.1);
        background: rgba(255,255,255,0.05);
        color: var(--bsh-text);
        outline: none;
      }

      .bsh-input:focus {
        border-color: rgba(79, 140, 255, 0.42);
        box-shadow: 0 0 0 3px rgba(79, 140, 255, 0.12);
      }

      .bsh-log {
        max-height: 140px;
        overflow-y: auto;
        padding: 10px 12px;
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.08);
        background: rgba(0,0,0,0.18);
      }

      .bsh-log__item {
        font-size: 11px;
        line-height: 1.5;
        color: var(--bsh-text-muted);
      }

      .bsh-log__item + .bsh-log__item {
        margin-top: 6px;
      }

      .bsh-log__item.is-success {
        color: #98e8bf;
      }

      .bsh-log__item.is-warning {
        color: #ffd692;
      }

      .bsh-log__item.is-error {
        color: #ffabab;
      }

      .bsh-inline-actions {
        margin-top: 10px;
      }

      .bsh-inline-actions .bsh-btn {
        padding: 7px 10px;
        font-size: 12px;
      }

      .bsh-toolbar {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-top: 12px;
        padding: 12px 14px;
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(79, 140, 255, 0.08), rgba(255,255,255,0.04));
        border: 1px solid rgba(79, 140, 255, 0.18);
      }

      .bsh-toolbar__title {
        font-size: 13px;
        font-weight: 700;
        color: #1f3c7a;
      }

      .bsh-toolbar__meta {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
      }

      .bsh-toolbar__status {
        flex-basis: 100%;
        font-size: 12px;
        color: rgba(20, 32, 64, 0.72);
      }

      .bsh-toolbar .bsh-badge {
        color: #24427f;
        background: rgba(79, 140, 255, 0.08);
        border-color: rgba(79, 140, 255, 0.16);
      }

      .bsh-toolbar .bsh-btn,
      .bsh-toolbar .bsh-btn:visited,
      .bsh-toolbar .bsh-btn:hover,
      .bsh-toolbar .bsh-btn:active {
        color: #1d3565 !important;
        background: rgba(255,255,255,0.9);
        border-color: rgba(79,140,255,0.16);
      }

      .bsh-toolbar .bsh-btn--primary,
      .bsh-toolbar .bsh-btn--primary:hover,
      .bsh-toolbar .bsh-btn--primary:active,
      .bsh-toolbar .bsh-btn--primary:visited {
        color: #fff !important;
      }

      .bsh-inline-actions .bsh-btn,
      .bsh-inline-actions .bsh-btn:visited,
      .bsh-inline-actions .bsh-btn:hover,
      .bsh-inline-actions .bsh-btn:active {
        color: #1d3565 !important;
        background: rgba(255,255,255,0.96);
        border-color: rgba(79,140,255,0.18);
      }

      .bsh-inline-actions .bsh-btn--primary,
      .bsh-inline-actions .bsh-btn--primary:hover,
      .bsh-inline-actions .bsh-btn--primary:active,
      .bsh-inline-actions .bsh-btn--primary:visited {
        color: #fff !important;
      }

      .bsh-inline-actions .bsh-btn:disabled,
      .bsh-inline-actions .bsh-btn:disabled:hover,
      .bsh-inline-actions .bsh-btn:disabled:active {
        color: rgba(29,53,101,0.6) !important;
        background: rgba(255,255,255,0.72);
      }

      .bsh-link-tools {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }

      .bsh-source {
        font-size: 11px;
        line-height: 1.5;
        color: var(--bsh-text-muted);
        word-break: break-all;
      }
    `;
    document.head.appendChild(style);
  }

  function applyCourseButtonStyle(button, { primary = false, disabled = false } = {}) {
    if (!button) return;
    button.style.setProperty("color", primary ? "#ffffff" : disabled ? "rgba(29,53,101,0.62)" : "#1d3565", "important");
    button.style.setProperty(
      "background",
      primary ? "linear-gradient(135deg, var(--bsh-primary), var(--bsh-primary-strong))" : disabled ? "rgba(255,255,255,0.72)" : "rgba(255,255,255,0.96)",
      "important"
    );
    button.style.setProperty("border-color", "rgba(79,140,255,0.18)", "important");
    button.style.setProperty("opacity", disabled ? "0.78" : "1", "important");
    button.style.setProperty("text-shadow", "none", "important");
    button.style.setProperty("box-shadow", "none", "important");
    button.style.setProperty("-webkit-text-fill-color", primary ? "#ffffff" : disabled ? "rgba(29,53,101,0.62)" : "#1d3565", "important");
  }

  function updateBadge(badge, ready, text, warning = false) {
    if (!badge) return;
    badge.textContent = text;
    badge.classList.toggle("is-ready", ready);
    badge.classList.toggle("is-warning", warning);
  }

  function initLivingroom() {
    installLivingroomInterceptor();

    const shell = createPanelShell(
      "bsh-livingroom-panel",
      "MSA 回放助手",
      "字幕、笔记与 PPT 讲义导出"
    );

    shell.body.innerHTML = `
      <div class="bsh-section">
        <div class="bsh-badges">
          <span class="bsh-badge" data-role="subtitle-badge">字幕未就绪</span>
          <span class="bsh-badge" data-role="ppt-badge">PPT 未就绪</span>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card bsh-card--status" data-role="status">
          等待页面请求数据。先点语音或 PPT 标签后，再回来导出。
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-actions">
          <button type="button" class="bsh-btn bsh-btn--primary" data-action="ppt" disabled>导出 PPT 讲义</button>
          <button type="button" class="bsh-btn" data-action="srt" disabled>导出 SRT 字幕</button>
          <button type="button" class="bsh-btn" data-action="txt" disabled>导出 TXT 笔记</button>
          <button type="button" class="bsh-btn bsh-btn--ghost" data-action="course-detail">当前课程下载页</button>
          <button type="button" class="bsh-btn bsh-btn--ghost" data-action="course-center">打开课程中心</button>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card">
          <ul class="bsh-hints">
            <li>导出字幕或笔记前，请先点开页面内的语音/字幕相关标签。</li>
            <li>导出 PPT 讲义前，请先点开页面内的 PPT 标签以触发数据请求。</li>
            <li>PPT 导出会打开打印页，浏览器允许弹窗后可另存为 PDF。</li>
            <li>如需进入课程下载页，可使用上方的“当前课程下载页”或“打开课程中心”按钮。</li>
            <li>如需搜索全校课程并进入可下载视频的课程详情页，可打开课程中心：<a href="https://classroom.msa.buaa.edu.cn/courseCenter" target="_blank" rel="noopener noreferrer">classroom.msa.buaa.edu.cn/courseCenter</a></li>
          </ul>
        </div>
      </div>
    `;

    state.livingroom.ui = {
      shell,
      subtitleBadge: shell.body.querySelector('[data-role="subtitle-badge"]'),
      pptBadge: shell.body.querySelector('[data-role="ppt-badge"]'),
      status: shell.body.querySelector('[data-role="status"]'),
      pptButton: shell.body.querySelector('[data-action="ppt"]'),
      srtButton: shell.body.querySelector('[data-action="srt"]'),
      txtButton: shell.body.querySelector('[data-action="txt"]'),
      courseDetailButton: shell.body.querySelector('[data-action="course-detail"]'),
      courseCenterButton: shell.body.querySelector('[data-action="course-center"]'),
    };

    state.livingroom.ui.pptButton.addEventListener("click", exportLivingroomPpt);
    state.livingroom.ui.srtButton.addEventListener("click", () => exportLivingroomSubtitle("srt"));
    state.livingroom.ui.txtButton.addEventListener("click", () => exportLivingroomSubtitle("txt"));
    state.livingroom.ui.courseDetailButton.addEventListener("click", () => {
      const courseDetailUrl = getCurrentCourseDetailUrl();
      if (!courseDetailUrl) {
        window.alert("当前页面没有 course_id，无法定位到对应课程下载页。");
        return;
      }
      window.open(courseDetailUrl, "_blank", "noopener,noreferrer");
    });
    state.livingroom.ui.courseCenterButton.addEventListener("click", () => {
      window.open(getCourseCenterUrl(), "_blank", "noopener,noreferrer");
    });

    updateLivingroomUi();
  }

  function installLivingroomInterceptor() {
    if (state.livingroom.interceptorInstalled) return;
    state.livingroom.interceptorInstalled = true;

    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (_method, url) {
      this.__bshUrl = String(url || "");
      return originalOpen.apply(this, arguments);
    };

    XMLHttpRequest.prototype.send = function () {
      this.addEventListener(
        "load",
        function () {
          const url = this.__bshUrl || "";
          if (!url) return;

          if (url.includes("search-trans-result")) {
            try {
              const response = JSON.parse(this.responseText || "{}");
              if (Array.isArray(response.list) && response.list.length > 0) {
                state.livingroom.subtitleData = response.list;
                updateLivingroomUi();
              }
            } catch (error) {
              console.warn("[BUAA Helper] 字幕数据解析失败", error);
            }
          }

          if (url.includes("search-ppt")) {
            try {
              const response = JSON.parse(this.responseText || "{}");
              if (Array.isArray(response.list) && response.list.length > 0) {
                state.livingroom.pptData = response.list;
                updateLivingroomUi();
              }
            } catch (error) {
              console.warn("[BUAA Helper] PPT 数据解析失败", error);
            }
          }
        },
        { once: true }
      );
      return originalSend.apply(this, arguments);
    };
  }

  function updateLivingroomUi() {
    const ui = state.livingroom.ui;
    if (!ui) return;

    const subtitleReady = Array.isArray(state.livingroom.subtitleData) && state.livingroom.subtitleData.length > 0;
    const pptReady = Array.isArray(state.livingroom.pptData) && state.livingroom.pptData.length > 0;

    updateBadge(ui.subtitleBadge, subtitleReady, subtitleReady ? "字幕已就绪" : "字幕未就绪");
    updateBadge(ui.pptBadge, pptReady, pptReady ? "PPT 已就绪" : "PPT 未就绪");

    ui.srtButton.disabled = !subtitleReady;
    ui.txtButton.disabled = !subtitleReady;
    ui.pptButton.disabled = !pptReady;

    if (subtitleReady && pptReady) {
      ui.status.textContent = "字幕与 PPT 数据都已准备完成，可以直接导出。";
    } else if (subtitleReady) {
      ui.status.textContent = "字幕数据已就绪。若要导出 PPT 讲义，请先点击页面里的 PPT 标签。";
    } else if (pptReady) {
      ui.status.textContent = "PPT 数据已就绪。若要导出字幕或笔记，请先点击页面里的语音/字幕标签。";
    } else {
      ui.status.textContent = "等待页面请求数据。先点语音或 PPT 标签后，再回来导出。";
    }
  }

  function getLivingroomRawLines() {
    return (state.livingroom.subtitleData || []).reduce((all, chapter) => {
      if (Array.isArray(chapter.all_content)) {
        all.push(...chapter.all_content);
      }
      return all;
    }, []);
  }

  function mergeSubtitleLines(rawLines) {
    if (!rawLines.length) return [];
    const merged = [];
    let current = { ...rawLines[0] };
    for (let index = 1; index < rawLines.length; index += 1) {
      const next = rawLines[index];
      const gap = Number(next.BeginSec || 0) - Number(current.EndSec || current.BeginSec || 0);
      const currentText = String(current.Text || "");
      if (gap < 1 && currentText.length < 20) {
        current.Text = `${currentText}，${next.Text || ""}`;
        current.EndSec = next.EndSec;
      } else {
        merged.push(current);
        current = { ...next };
      }
    }
    merged.push(current);
    return merged;
  }

  function exportLivingroomSubtitle(type) {
    const rawLines = getLivingroomRawLines();
    if (!rawLines.length) {
      window.alert("还没有捕获到字幕数据，请先点击页面里的语音或字幕标签。");
      return;
    }

    const lines = type === "txt" ? mergeSubtitleLines(rawLines) : rawLines;
    const safeTitle = sanitizeFilename(document.title, 40);

    let content = type === "txt" ? `课程笔记 - ${document.title}\n\n` : "";
    lines.forEach((line, index) => {
      const begin = Number(line.BeginSec || 0);
      const end = Number(line.EndSec || begin + 2);
      const text = String(line.Text || "").trim();
      if (!text) return;
      if (type === "srt") {
        content += `${index + 1}\n${formatSrtTime(begin)} --> ${formatSrtTime(end)}\n${text}\n\n`;
      } else {
        content += `[${formatClockTime(begin)}] ${text}\n`;
      }
    });

    const fileName = `${safeTitle}_${type === "srt" ? "字幕" : "笔记"}.${type}`;
    downloadTextFile(content, fileName);
  }

  function exportLivingroomPpt() {
    const pptData = Array.isArray(state.livingroom.pptData) ? state.livingroom.pptData : [];
    if (!pptData.length) {
      window.alert("还没有捕获到 PPT 数据，请先点击页面里的 PPT 标签。");
      return;
    }

    const slides = pptData
      .map((slide, index) => {
        try {
          const content = slide.content ? JSON.parse(slide.content) : {};
          const imageUrl = content.pptimgurl;
          if (!imageUrl) return null;
          const timeText = formatClockTime(Number(slide.created_sec || slide.BeginSec || 0));
          return { imageUrl, index: index + 1, timeText };
        } catch (error) {
          return null;
        }
      })
      .filter(Boolean);

    const title = escapeHtml(document.title);
    const generatedAt = escapeHtml(new Date().toLocaleString("zh-CN"));

    const slidesMarkup = slides
      .map(
        (slide) => `
          <section class="slide-card">
            <img class="ppt-image" src="${escapeHtml(slide.imageUrl)}" alt="第 ${slide.index} 页">
            <div class="slide-meta">
              <span>第 ${slide.index} 页</span>
              <span>对应时间：${escapeHtml(slide.timeText)}</span>
            </div>
          </section>
        `
      )
      .join("");

    const html = `
      <html lang="zh-CN">
        <head>
          <meta charset="UTF-8">
          <title>${title} - PPT 讲义</title>
          <style>
            * { box-sizing: border-box; }
            body {
              margin: 0;
              padding: 24px;
              background: #eef3fb;
              font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
              color: #0f172a;
            }
            .loading-mask {
              position: fixed;
              inset: 0;
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: center;
              gap: 10px;
              background: rgba(255,255,255,0.94);
              z-index: 9999;
            }
            .loading-mask__title { font-size: 24px; font-weight: 700; color: #1d4ed8; }
            .loading-mask__desc { font-size: 14px; color: #475569; }
            .doc-header {
              max-width: 1024px;
              margin: 0 auto 20px;
              padding: 20px 22px;
              border-radius: 18px;
              background: white;
              box-shadow: 0 14px 34px rgba(15, 23, 42, 0.08);
            }
            .doc-header h1 { margin: 0 0 8px; font-size: 28px; }
            .doc-header p { margin: 6px 0 0; color: #475569; font-size: 14px; }
            .hint {
              margin-top: 10px;
              padding: 10px 12px;
              border-radius: 12px;
              background: #eff6ff;
              color: #1d4ed8;
              font-size: 13px;
            }
            .slide-card {
              width: min(1024px, 100%);
              margin: 0 auto 24px;
              overflow: hidden;
              border-radius: 18px;
              border: 1px solid rgba(148, 163, 184, 0.28);
              background: white;
              box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
            }
            .ppt-image {
              display: block;
              width: 100%;
              background: white;
            }
            .slide-meta {
              display: flex;
              justify-content: space-between;
              gap: 12px;
              padding: 12px 16px;
              color: #475569;
              font-size: 13px;
              border-top: 1px solid rgba(148, 163, 184, 0.22);
            }
            @media print {
              @page { margin: 1cm; }
              body { background: white; padding: 0; }
              .loading-mask, .hint { display: none !important; }
              .doc-header {
                padding: 0 0 12px;
                border-radius: 0;
                box-shadow: none;
                border-bottom: 2px solid #334155;
              }
              .doc-header h1 { font-size: 18pt; }
              .doc-header p { font-size: 10pt; }
              .slide-card {
                width: 100%;
                margin: 0;
                border-radius: 0;
                box-shadow: none;
                page-break-after: always;
                border: 1px solid #94a3b8;
              }
              .ppt-image {
                max-width: 100%;
                max-height: 82vh;
                width: auto;
                margin: 0 auto;
              }
              .slide-meta {
                font-size: 10pt;
                color: #334155;
                background: #f8fafc !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
              }
            }
          </style>
        </head>
        <body tabindex="-1">
          <div class="loading-mask" id="loading-mask">
            <div class="loading-mask__title">正在准备 PPT 讲义</div>
            <div class="loading-mask__desc" id="loading-text">正在预加载图片，稍后会自动唤起打印。</div>
          </div>

          <header class="doc-header">
            <h1>${title}</h1>
            <p>生成时间：${generatedAt}</p>
            <p>导出方式：浏览器打印页，可在打印窗口中选择“另存为 PDF”。</p>
            <div class="hint">建议在打印设置中手动选择横向布局，以获得更舒适的讲义排版。</div>
          </header>

          ${slidesMarkup || '<div class="doc-header"><p>未找到可导出的 PPT 图片。</p></div>'}

          <script>
            window.onload = function () {
              const images = document.querySelectorAll('.ppt-image');
              const loadingText = document.getElementById('loading-text');
              const mask = document.getElementById('loading-mask');
              let loaded = 0;
              const total = images.length;

              function step() {
                loaded += 1;
                loadingText.textContent = '正在预加载图片（' + loaded + '/' + total + '）';
                if (loaded >= total) {
                  loadingText.textContent = '资源准备完成，正在唤起打印。';
                  setTimeout(function () {
                    mask.style.display = 'none';
                    window.print();
                  }, 450);
                }
              }

              if (total === 0) {
                step();
                return;
              }

              images.forEach(function (image) {
                if (image.complete) {
                  step();
                } else {
                  image.onload = step;
                  image.onerror = step;
                }
              });
            };
          <\/script>
        </body>
      </html>
    `;

    const popup = openHtmlInPopup(html);
    if (!popup) {
      window.alert("浏览器拦截了弹窗。请允许当前站点弹窗后，再尝试导出 PPT 讲义。");
    }
  }

  async function initCourseDetail() {
    const courseId = state.courseDetail.courseId;
    if (!courseId) return;

    state.courseDetail.token = getJwtTokenFromCookieToken();
    state.courseDetail.account = getJwtAccountFromCookieToken();

    const rows = await waitFor(() => {
      const found = Array.from(document.querySelectorAll("div.content-inner-one > p"));
      return found.length ? found : null;
    });
    if (!rows) return;

    state.courseDetail.rowCount = rows.length;

    ensureCourseToolbar();

    if (!state.courseDetail.token) {
      updateCourseToolbarStatus("未检测到登录凭证，课程详情页下载功能不可用。", "warning");
      return;
    }

    try {
      state.courseDetail.courseData = await courseApiGet(
        `https://yjapi.msa.buaa.edu.cn/courseapi/v3/multi-search/get-course-detail?course_id=${encodeURIComponent(courseId)}${state.courseDetail.account ? `&student=${encodeURIComponent(state.courseDetail.account)}` : ""}`
      ).then((response) => response.data);

      state.courseDetail.sessions = flattenCourseSessions(state.courseDetail.courseData);
      bindCourseDetailRows(rows, state.courseDetail.sessions);
      updateCourseToolbarMeta();
      updateCourseToolbarStatus("课程详情工具已就绪，可按节下载视频或导出 PPTX。", "success");
    } catch (error) {
      console.error("[BUAA Helper] 课程详情初始化失败", error);
      updateCourseToolbarStatus("课程详情数据加载失败，请刷新页面后重试。", "warning");
    }
  }

  function ensureCourseToolbar() {
    let host = document.querySelector("div.content-tips");
    if (!host) {
      host = document.body;
    }

    const existing = document.getElementById("bsh-course-toolbar");
    if (existing) existing.remove();

    const toolbar = document.createElement("section");
    toolbar.id = "bsh-course-toolbar";
    toolbar.className = "bsh-toolbar";
    toolbar.innerHTML = `
      <div>
        <div class="bsh-toolbar__title">课程详情工具</div>
        <div class="bsh-toolbar__meta">
          <span class="bsh-badge" data-role="count">等待加载</span>
          <span class="bsh-badge" data-role="available">待检测</span>
        </div>
      </div>
      <div class="bsh-toolbar__actions">
        <button type="button" class="bsh-btn bsh-btn--primary" data-action="all-video" disabled>下载全部视频</button>
        <button type="button" class="bsh-btn" data-action="all-ppt" disabled>下载全部 PPT</button>
      </div>
      <div class="bsh-toolbar__status" data-role="status">正在准备课程详情工具…</div>
    `;

    host.appendChild(toolbar);

    state.courseDetail.ui = {
      toolbar,
      countBadge: toolbar.querySelector('[data-role="count"]'),
      availableBadge: toolbar.querySelector('[data-role="available"]'),
      status: toolbar.querySelector('[data-role="status"]'),
      allVideoButton: toolbar.querySelector('[data-action="all-video"]'),
      allPptButton: toolbar.querySelector('[data-action="all-ppt"]'),
    };

    applyCourseButtonStyle(state.courseDetail.ui.allVideoButton, { primary: true, disabled: true });
    applyCourseButtonStyle(state.courseDetail.ui.allPptButton, { disabled: true });

    state.courseDetail.ui.allVideoButton.addEventListener("click", runBulkVideoDownload);
    state.courseDetail.ui.allPptButton.addEventListener("click", runBulkPptExport);
  }

  function updateCourseToolbarStatus(text, tone = "info") {
    const ui = state.courseDetail.ui;
    if (!ui) return;
    ui.status.textContent = text;
    ui.status.dataset.tone = tone;
  }

  function updateCourseToolbarMeta() {
    const ui = state.courseDetail.ui;
    if (!ui) return;
    const total = state.courseDetail.sessions.length || state.courseDetail.rowCount || 0;
    const available = state.courseDetail.sessions.length
      ? state.courseDetail.sessions.filter((session) => String(session.sub_status || session.status) === COURSE_READY_STATUS).length
      : Array.from(document.querySelectorAll('.bsh-inline-actions [data-action="download"]')).filter(
          (button) => !button.disabled
        ).length;
    ui.countBadge.textContent = `共 ${total} 节`;
    ui.availableBadge.textContent = `${available} 节可下载`;
    ui.allVideoButton.disabled = available === 0;
    ui.allPptButton.disabled = available === 0;
    applyCourseButtonStyle(ui.allVideoButton, { primary: true, disabled: ui.allVideoButton.disabled });
    applyCourseButtonStyle(ui.allPptButton, { disabled: ui.allPptButton.disabled });
  }

  async function runBulkVideoDownload() {
    const buttons = Array.from(document.querySelectorAll('.bsh-inline-actions [data-action="download"]')).filter(
      (button) => !button.disabled
    );
    if (!buttons.length) return;

    setBusy(state.courseDetail.ui.allVideoButton, true, "下载中");
    updateCourseToolbarStatus("正在批量触发视频下载，浏览器可能连续弹出多个下载。", "info");
    try {
      for (const button of buttons) {
        button.click();
        await sleep(900);
      }
      updateCourseToolbarStatus("已完成批量视频下载触发。", "success");
    } finally {
      setBusy(state.courseDetail.ui.allVideoButton, false);
    }
  }

  async function runBulkPptExport() {
    const buttons = Array.from(document.querySelectorAll('.bsh-inline-actions [data-action="pptx"]')).filter(
      (button) => !button.disabled
    );
    if (!buttons.length) return;

    setBusy(state.courseDetail.ui.allPptButton, true, "导出中");
    updateCourseToolbarStatus("正在按顺序导出 PPTX，请等待浏览器完成文件写出。", "info");
    try {
      for (const button of buttons) {
        button.click();
        await sleep(900);
      }
      updateCourseToolbarStatus("已完成批量 PPTX 导出触发。", "success");
    } finally {
      setBusy(state.courseDetail.ui.allPptButton, false);
    }
  }

  function flattenCourseSessions(courseData) {
    const sessions = [];
    const subList = courseData?.sub_list || courseData?.data?.sub_list || {};

    const visit = (value) => {
      if (!value) return;
      if (Array.isArray(value)) {
        value.forEach(visit);
        return;
      }
      if (typeof value !== "object") return;
      if (
        ("id" in value || "sub_id" in value) &&
        ("sub_title" in value || "title" in value || "lecturer_name" in value || "status" in value || "sub_status" in value)
      ) {
        sessions.push(value);
        return;
      }
      Object.values(value).forEach(visit);
    };

    visit(subList);
    return sessions;
  }

  function bindCourseDetailRows(rows, sessions) {
    const mapping = mapRowsToSessions(rows, sessions);
    const courseName = document.querySelector("div.info-title > p")?.textContent?.trim() || "课程";

    mapping.forEach(({ row, session, matchedBy }, index) => {
      if (!row || !session) return;
      const previous = row.querySelector(".bsh-inline-actions");
      if (previous) previous.remove();

      const actions = document.createElement("div");
      actions.className = "bsh-inline-actions";

      const isAvailable = String(session.sub_status || session.status) === COURSE_READY_STATUS;
      const sessionId = session.id || session.sub_id;
      const displayName = sanitizeFilename(`${courseName}${session.sub_title || `第${index + 1}节`}-${session.lecturer_name || "授课教师"}`);

      const downloadButton = document.createElement("button");
      downloadButton.type = "button";
      downloadButton.className = "bsh-btn bsh-btn--primary";
      downloadButton.dataset.action = "download";
      downloadButton.textContent = "下载视频";
      downloadButton.disabled = !isAvailable;
      applyCourseButtonStyle(downloadButton, { primary: true, disabled: downloadButton.disabled });
      downloadButton.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        await handleCourseVideoDownload(sessionId, displayName, downloadButton);
      });

      const previewButton = document.createElement("button");
      previewButton.type = "button";
      previewButton.className = "bsh-btn bsh-btn--ghost";
      previewButton.dataset.action = "preview";
      previewButton.textContent = "预览视频";
      previewButton.disabled = !isAvailable;
      applyCourseButtonStyle(previewButton, { disabled: previewButton.disabled });
      previewButton.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        await handleCourseVideoPreview(sessionId, displayName, previewButton);
      });

      const pptButton = document.createElement("button");
      pptButton.type = "button";
      pptButton.className = "bsh-btn";
      pptButton.dataset.action = "pptx";
      pptButton.textContent = "导出 PPTX";
      pptButton.disabled = !isAvailable;
      applyCourseButtonStyle(pptButton, { disabled: pptButton.disabled });
      pptButton.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        await handleCoursePptExport(sessionId, pptButton);
      });

      actions.appendChild(downloadButton);
      actions.appendChild(previewButton);
      actions.appendChild(pptButton);
      row.appendChild(actions);

      if (rows.length !== sessions.length && matchedBy === "fallback") {
        actions.title = "当前行通过顺序回退匹配到课程数据，如按钮对应不准确，请刷新页面重试。";
      }
    });

    if (rows.length !== sessions.length) {
      updateCourseToolbarStatus(
        `页面行数与课程数据数量不完全一致（页面 ${rows.length}，数据 ${sessions.length}），已尽量按标题和顺序匹配。`,
        "warning"
      );
    }
  }

  function mapRowsToSessions(rows, sessions) {
    const normalizedSessions = sessions.map((session) => normalizeText(session.sub_title));
    const used = new Set();
    const mapped = rows.map((row) => {
      const rowText = normalizeText(row.textContent || "");
      const matchedIndex = normalizedSessions.findIndex((title, index) => {
        if (used.has(index) || !title) return false;
        return rowText.includes(title) || title.includes(rowText);
      });

      if (matchedIndex >= 0) {
        used.add(matchedIndex);
        return { row, session: sessions[matchedIndex], matchedBy: "title" };
      }

      return { row, session: null, matchedBy: "fallback" };
    });

    let fallbackIndex = 0;
    mapped.forEach((item) => {
      if (item.session) return;
      while (used.has(fallbackIndex) && fallbackIndex < sessions.length) {
        fallbackIndex += 1;
      }
      if (fallbackIndex < sessions.length) {
        item.session = sessions[fallbackIndex];
        used.add(fallbackIndex);
      }
      fallbackIndex += 1;
    });

    return mapped;
  }

  async function courseApiGet(url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${state.courseDetail.token}`,
      },
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return response.json();
  }

  async function getCourseUserInfo() {
    if (!state.courseDetail.userInfoPromise) {
      state.courseDetail.userInfoPromise = courseApiGet(
        "https://classroom.msa.buaa.edu.cn/userapi/v1/infosimple"
      );
    }
    return state.courseDetail.userInfoPromise;
  }

  async function getCourseSubInfo(subId) {
    const courseId = state.courseDetail.courseId;
    return courseApiGet(
      `https://classroom.msa.buaa.edu.cn/courseapi/v3/portal-home-setting/get-sub-info?course_id=${encodeURIComponent(
        courseId
      )}&sub_id=${encodeURIComponent(subId)}`
    ).then((response) => response.data);
  }

  async function buildSignedVideoUrl(rawUrl) {
    const userInfo = await getCourseUserInfo();
    const user = userInfo?.params || {};
    if (!user.phone || !user.tenant_id || !user.id) {
      throw new Error("用户信息不完整，无法生成签名视频地址。");
    }

    const target = new URL(rawUrl, window.location.origin);
    target.searchParams.set("clientUUID", generateUuid());

    const epoch = Math.floor(Date.now() / 1000);
    const reversedPhone = String(user.phone).split("").reverse().join("");
    const hash = md5Hex(`${target.pathname}${user.id}${user.tenant_id}${reversedPhone}${epoch}`);
    target.searchParams.set("t", `${user.id}-${epoch}-${hash}`);
    return target.toString();
  }

  async function getSignedVideoStreams(subId) {
    const subInfo = await getCourseSubInfo(subId);
    const list = Object.values(subInfo?.video_list || {});
    const streams = [];

    for (const video of list) {
      if (!video?.preview_url) continue;
      const signedUrl = await buildSignedVideoUrl(video.preview_url);
      streams.push({
        url: signedUrl,
        label: inferStreamLabel(video.preview_url),
      });
      await sleep(200);
    }

    return streams;
  }

  function inferStreamLabel(url) {
    if (url.includes("ppt")) return "屏幕";
    if (url.includes("tea")) return "黑板";
    return "视频";
  }

  async function handleCourseVideoDownload(subId, baseName, button) {
    setBusy(button, true, "下载中");
    updateCourseToolbarStatus(`正在准备 ${baseName} 的视频下载。`, "info");
    try {
      const streams = await getSignedVideoStreams(subId);
      if (!streams.length) {
        updateCourseToolbarStatus(`未找到 ${baseName} 的可下载视频地址。`, "warning");
        return;
      }
      for (const stream of streams) {
        triggerUrlDownload(stream.url, `${sanitizeFilename(`${baseName}-${stream.label}`)}.mp4`);
        await sleep(350);
      }
      updateCourseToolbarStatus(`${baseName} 的视频下载已触发。`, "success");
    } catch (error) {
      console.error(error);
      updateCourseToolbarStatus(`${baseName} 的视频下载失败，请稍后重试。`, "warning");
    } finally {
      setBusy(button, false);
    }
  }

  async function handleCourseVideoPreview(subId, baseName, button) {
    setBusy(button, true, "准备中");
    updateCourseToolbarStatus(`正在生成 ${baseName} 的预览页面。`, "info");
    try {
      const streams = await getSignedVideoStreams(subId);
      if (!streams.length) {
        updateCourseToolbarStatus(`未找到 ${baseName} 的可预览视频地址。`, "warning");
        return;
      }
      openVideoPreviewWindow(baseName, streams);
      updateCourseToolbarStatus(`${baseName} 的预览页面已打开。`, "success");
    } catch (error) {
      console.error(error);
      updateCourseToolbarStatus(`${baseName} 的预览页面打开失败。`, "warning");
    } finally {
      setBusy(button, false);
    }
  }

  function openVideoPreviewWindow(title, streams) {
    const html = `
      <html lang="zh-CN">
        <head>
          <meta charset="UTF-8">
          <title>${escapeHtml(title)} - 视频预览</title>
          <style>
            body {
              margin: 0;
              padding: 28px;
              background: #0f172a;
              color: #f8fafc;
              font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            }
            .page-title {
              font-size: 28px;
              font-weight: 700;
              margin-bottom: 18px;
            }
            .stream-card {
              margin-bottom: 22px;
              padding: 18px;
              border-radius: 18px;
              background: rgba(255,255,255,0.05);
              border: 1px solid rgba(255,255,255,0.08);
            }
            .stream-title {
              margin-bottom: 12px;
              font-size: 18px;
              font-weight: 600;
            }
            video {
              width: 100%;
              max-width: 1200px;
              border-radius: 14px;
              background: black;
            }
            .stream-url {
              margin-top: 10px;
              color: rgba(248,250,252,0.75);
              font-size: 12px;
              word-break: break-all;
            }
            .shortcut-hint {
              position: fixed;
              bottom: 24px;
              right: 24px;
              padding: 10px 16px;
              border-radius: 12px;
              background: rgba(255,255,255,0.08);
              border: 1px solid rgba(255,255,255,0.12);
              font-size: 13px;
              color: rgba(248,250,252,0.6);
              pointer-events: none;
            }
          </style>
        </head>
        <body>
          <div class="page-title">${escapeHtml(title)}</div>
          ${streams
            .map(
              (stream) => `
                <section class="stream-card">
                  <div class="stream-title">${escapeHtml(stream.label)}</div>
                  <video controls preload="metadata" src="${escapeHtml(stream.url)}"></video>
                  <div class="stream-url">${escapeHtml(stream.url)}</div>
                </section>
              `
            )
            .join("")}
          <div class="shortcut-hint">快捷键：← 后退 10s | → 快进 10s</div>
          <script>
            (function () {
              let activeVideo = null;

              const pickVideo = () => {
                const videos = Array.from(document.querySelectorAll('video'));
                if (!videos.length) return null;
                if (activeVideo && videos.includes(activeVideo)) return activeVideo;
                return videos.find((v) => !v.paused) || videos[0];
              };

              const seekBy = (offset) => {
                const video = pickVideo();
                if (!video) return;
                const maxDuration = Number.isFinite(video.duration) ? video.duration : Number.MAX_SAFE_INTEGER;
                const next = Math.max(0, Math.min(maxDuration, (video.currentTime || 0) + offset));
                video.currentTime = next;
              };

              document.querySelectorAll('video').forEach((video) => {
                video.addEventListener('pointerdown', () => {
                  activeVideo = video;
                });
                video.addEventListener('play', () => {
                  activeVideo = video;
                });
              });

              const onKeyDown = (event) => {
                const tag = (event.target && event.target.tagName) ? event.target.tagName.toUpperCase() : '';
                if (tag === 'INPUT' || tag === 'TEXTAREA' || (event.target && event.target.isContentEditable)) {
                  return;
                }

                if (event.key === 'ArrowRight') {
                  event.preventDefault();
                  event.stopPropagation();
                  seekBy(10);
                } else if (event.key === 'ArrowLeft') {
                  event.preventDefault();
                  event.stopPropagation();
                  seekBy(-10);
                }
              };

              document.addEventListener('keydown', onKeyDown, true);
              window.addEventListener('keydown', onKeyDown, true);

              window.addEventListener('load', () => {
                try {
                  document.body.focus();
                } catch (error) {
                  // ignore
                }
              });
            })();
          </script>
        </body>
      </html>
    `;

    const popup = openHtmlInPopup(html);
    if (!popup) {
      throw new Error("浏览器拦截了视频预览弹窗。");
    }
  }

  async function handleCoursePptExport(subId, button) {
    setBusy(button, true, "导出中");
    updateCourseToolbarStatus("正在导出 PPTX，请等待浏览器完成文件写出。", "info");
    try {
      const subInfo = await getCourseSubInfo(subId);
      const pptResponse = await courseApiGet(
        `https://classroom.msa.buaa.edu.cn/pptnote/v1/schedule/search-ppt?course_id=${encodeURIComponent(
          state.courseDetail.courseId
        )}&sub_id=${encodeURIComponent(subId)}&resource_guid=${encodeURIComponent(subInfo.resource_guid || "")}`
      );
      const imageUrls = (pptResponse.list || [])
        .map((item) => {
          try {
            return JSON.parse(item.content || "{}").pptimgurl;
          } catch (error) {
            return "";
          }
        })
        .filter(Boolean);

      if (!imageUrls.length) {
        updateCourseToolbarStatus("当前课程未找到可导出的 PPT 图片。", "warning");
        return;
      }

      if (typeof PptxGenJS !== "function") {
        throw new Error("PptxGenJS 未加载成功。");
      }

      const title = sanitizeFilename(`${subInfo.course_title || "课程"}${subInfo.sub_title || "课时"}`);
      const ppt = new PptxGenJS();
      ppt.author = subInfo.lecturer_name || "micraow";
      ppt.title = title;
      ppt.subject = subInfo.course_title || "BUAA Course";
      ppt.company = "BUAA";
      ppt.layout = "LAYOUT_16x9";
      imageUrls.forEach((imageUrl) => {
        const slide = ppt.addSlide();
        slide.background = { path: imageUrl };
      });
      await ppt.writeFile({ fileName: `${title}.pptx` });
      updateCourseToolbarStatus(`${title}.pptx 导出完成。`, "success");
    } catch (error) {
      console.error(error);
      updateCourseToolbarStatus("PPTX 导出失败，请稍后重试。", "warning");
    } finally {
      setBusy(button, false);
    }
  }

  function ensureSpocPlaybackRate(force = false) {
    const video = state.spoc.currentVideo;
    const targetRate = Number(state.spoc.preferredRate);
    if (!video || !Number.isFinite(targetRate) || targetRate < 0.25 || targetRate > 16) {
      return;
    }

    if (force || Math.abs((video.playbackRate || 1) - targetRate) > 0.001) {
      video.playbackRate = targetRate;
    }
  }

  function startSpocRateEnforcer() {
    if (state.spoc.rateEnforceTimer) {
      clearInterval(state.spoc.rateEnforceTimer);
    }
    state.spoc.rateEnforceTimer = window.setInterval(() => {
      ensureSpocPlaybackRate(false);
    }, 500);
  }

  function installSpocRequestInterceptor() {
    if (state.spoc.interceptorInstalled) return;
    state.spoc.interceptorInstalled = true;

    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;
    const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
      this.__bshSpocUrl = String(url || "");
      this.__bshSpocHeaders = {};
      return originalOpen.apply(this, [method, url, ...rest]);
    };

    XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
      if (!this.__bshSpocHeaders) {
        this.__bshSpocHeaders = {};
      }
      this.__bshSpocHeaders[String(name || "")] = String(value || "");
      return originalSetRequestHeader.apply(this, [name, value]);
    };

    XMLHttpRequest.prototype.send = function (body) {
      const url = String(this.__bshSpocUrl || "");
      if (url.includes("addKcnrSfydNew")) {
        const params = parseSpocProgressRequestBody(body);
        if (params) {
          Promise.resolve(handleSpocCapturedRequest(params, this.__bshSpocHeaders || {})).catch(
            (error) => console.error(error)
          );
        } else {
          pushSpocLog("检测到 addKcnrSfydNew 请求，但参数解析失败。", "warning");
        }
      }
      return originalSend.apply(this, [body]);
    };

    pushSpocLog("SPOC 请求拦截已启用。", "success");
  }

  function parseSpocProgressRequestBody(body) {
    if (typeof body !== "string") return null;
    try {
      const payload = JSON.parse(body);
      const params = {
        kcnrid: payload?.kcnrid,
        kcid: payload?.kcid,
        ssmlid: payload?.ssmlid,
      };
      if (!params.kcnrid || !params.kcid || !params.ssmlid) return null;
      return params;
    } catch (error) {
      return null;
    }
  }

  function getSpocHeaderValue(name) {
    const lowerName = String(name || "").toLowerCase();
    const headers = state.spoc.capturedHeaders || {};
    const matchedKey = Object.keys(headers).find((key) => key.toLowerCase() === lowerName);
    return matchedKey ? String(headers[matchedKey] || "") : "";
  }

  async function handleSpocCapturedRequest(params, headers) {
    const prev = state.spoc.params;
    const changed =
      !prev ||
      prev.kcnrid !== params.kcnrid ||
      prev.kcid !== params.kcid ||
      prev.ssmlid !== params.ssmlid;

    state.spoc.params = {
      kcnrid: params.kcnrid,
      kcid: params.kcid,
      ssmlid: params.ssmlid,
    };
    state.spoc.capturedHeaders = {
      ...state.spoc.capturedHeaders,
      ...(headers || {}),
    };

    if (state.spoc.isRunning) {
      stopSpocSimulation("检测到新资源，已停止当前任务。");
    }

    if (changed) {
      pushSpocLog(`已捕获参数：KCID=${params.kcid}, KCNRID=${params.kcnrid}。`, "success");
    }

    const video = findBestVideoElement();
    const duration = await detectVideoDuration(video);
    if (duration > 0) {
      state.spoc.simDuration = duration;
      if (state.spoc.ui?.simDurationInput) {
        const current = Number(state.spoc.ui.simDurationInput.value || 0);
        if (!Number.isFinite(current) || current <= 0) {
          state.spoc.ui.simDurationInput.value = String(duration);
        }
      }
      pushSpocLog(`自动识别视频时长：${duration} 秒。`, "info");
    } else {
      pushSpocLog("未能自动识别时长，请手动填写。", "warning");
    }

    renderSpocUi();
  }

  function readSpocNumberInput(input, fallback, min, max) {
    const raw = Number(input?.value);
    let value = Number.isFinite(raw) ? raw : fallback;
    if (Number.isFinite(min)) value = Math.max(min, value);
    if (Number.isFinite(max)) value = Math.min(max, value);
    return value;
  }

  function syncSpocAutomationConfigFromUi() {
    const ui = state.spoc.ui;
    if (!ui) return;

    state.spoc.simSpeed = readSpocNumberInput(ui.simSpeedInput, state.spoc.simSpeed || 10, 0.25, 100);
    state.spoc.simInterval = readSpocNumberInput(ui.simIntervalInput, state.spoc.simInterval || 1, 0.5, 30);
    state.spoc.simDuration = readSpocNumberInput(ui.simDurationInput, state.spoc.simDuration || 0, 0, 43200);
    state.spoc.fastMode = Boolean(ui.fastModeInput?.checked);

    ui.simSpeedInput.value = String(state.spoc.simSpeed);
    ui.simIntervalInput.value = String(state.spoc.simInterval);
    if (state.spoc.simDuration > 0) {
      ui.simDurationInput.value = String(Math.floor(state.spoc.simDuration));
    }
  }

  function clearSpocSimulationTimer() {
    if (state.spoc.timer) {
      clearInterval(state.spoc.timer);
      state.spoc.timer = null;
    }
  }

  function stopSpocSimulation(statusText = "已停止。") {
    const wasRunning = state.spoc.isRunning;
    clearSpocSimulationTimer();
    state.spoc.isRunning = false;
    state.spoc.simTickBusy = false;
    state.spoc.simStatusText = statusText;
    if (wasRunning) {
      pushSpocLog(statusText, "info");
    }
    renderSpocUi();
  }

  function getSpocEffectiveDuration() {
    if (Number.isFinite(state.spoc.simDuration) && state.spoc.simDuration > 0) {
      return Math.floor(state.spoc.simDuration);
    }
    if (Number.isFinite(state.spoc.currentDuration) && state.spoc.currentDuration > 0) {
      return Math.floor(state.spoc.currentDuration);
    }
    return 0;
  }

  async function sendSpocApiRequest(endpoint, data) {
    return new Promise((resolve) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", `https://spoc.buaa.edu.cn/spocnewht${endpoint}`, true);
      xhr.withCredentials = true;

      xhr.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
      xhr.setRequestHeader("Accept", "application/json, text/plain, */*");

      const token = getSpocHeaderValue("Token");
      const roleCode = getSpocHeaderValue("RoleCode");
      const xRequestedWith = getSpocHeaderValue("X-Requested-With");
      const authorization = getSpocHeaderValue("Authorization");

      if (token) xhr.setRequestHeader("Token", token);
      if (roleCode) xhr.setRequestHeader("RoleCode", roleCode);
      if (xRequestedWith) xhr.setRequestHeader("X-Requested-With", xRequestedWith);
      if (authorization) xhr.setRequestHeader("Authorization", authorization);

      xhr.onreadystatechange = () => {
        if (xhr.readyState !== 4) return;
        resolve(xhr.status >= 200 && xhr.status < 300);
      };

      xhr.onerror = () => resolve(false);
      xhr.send(JSON.stringify(data || {}));
    });
  }

  async function addSpocStudyHeartbeat() {
    if (!state.spoc.params) return false;
    return sendSpocApiRequest("/kcnr/addNrydjlb", {
      kcnrid: state.spoc.params.kcnrid,
      kcid: state.spoc.params.kcid,
      nrlx: "99",
    });
  }

  async function updateSpocProgress(progress, playTime) {
    if (!state.spoc.params) return false;
    return sendSpocApiRequest("/kcnr/updKcnrSfydNew", {
      bfjd: progress,
      kcnrid: state.spoc.params.kcnrid,
      kcid: state.spoc.params.kcid,
      sfyd: progress >= 100 ? "1" : "0",
      bfsj: Math.floor(playTime),
      ssmlid: state.spoc.params.ssmlid,
    });
  }

  async function runSpocFastComplete(duration) {
    state.spoc.isRunning = true;
    state.spoc.simTickBusy = true;
    state.spoc.simPlayTime = 0;
    state.spoc.simProgress = 0;
    state.spoc.simDuration = duration;
    state.spoc.simStatusText = "正在执行一键完成...";
    state.spoc.simInfoText = "一键完成会直接将进度上报为 100%。";
    renderSpocUi();

    pushSpocLog("开始执行一键完成。", "warning");
    await addSpocStudyHeartbeat();
    const success = await updateSpocProgress(100, duration);

    state.spoc.simTickBusy = false;
    state.spoc.isRunning = false;

    if (success) {
      state.spoc.simProgress = 100;
      state.spoc.simPlayTime = duration;
      state.spoc.simStatusText = "一键完成执行成功。";
      state.spoc.simInfoText = "当前资源已标记为完成。";
      pushSpocLog("一键完成成功。", "success");
    } else {
      state.spoc.simStatusText = "一键完成执行失败。";
      state.spoc.simInfoText = "请重新捕获参数后重试。";
      pushSpocLog("一键完成请求失败。", "error");
    }
    renderSpocUi();
  }

  async function startSpocSimulation() {
    if (state.spoc.isRunning) {
      pushSpocLog("已有任务在运行。", "warning");
      return;
    }

    syncSpocAutomationConfigFromUi();

    if (!state.spoc.params) {
      state.spoc.simStatusText = "参数未就绪。";
      state.spoc.simInfoText = "请先点击资源“查看”以捕获请求参数。";
      renderSpocUi();
      pushSpocLog("缺少参数，请先打开一个资源。", "warning");
      return;
    }

    if (!getSpocHeaderValue("Token")) {
      state.spoc.simStatusText = "认证头未捕获。";
      state.spoc.simInfoText = "请重新点击“查看”以捕获 Token。";
      renderSpocUi();
      pushSpocLog("缺少 Token 请求头。", "error");
      return;
    }

    const duration = getSpocEffectiveDuration();
    if (!duration) {
      state.spoc.simStatusText = "时长无效。";
      state.spoc.simInfoText = "请手动填写时长或先加载视频元数据。";
      renderSpocUi();
      pushSpocLog("视频时长无效。", "warning");
      return;
    }

    state.spoc.simDuration = duration;
    if (state.spoc.fastMode) {
      await runSpocFastComplete(duration);
      return;
    }

    state.spoc.isRunning = true;
    state.spoc.simTickBusy = false;
    state.spoc.simPlayTime = 0;
    state.spoc.simProgress = 0;
    state.spoc.simStatusText = "正在模拟上报进度...";
    state.spoc.simInfoText = `速度 ${state.spoc.simSpeed}x | 心跳 ${state.spoc.simInterval}s`;
    renderSpocUi();

    pushSpocLog(
      `模拟开始：时长 ${duration}s，速度 ${state.spoc.simSpeed}x，心跳 ${state.spoc.simInterval}s。`,
      "success"
    );

    const heartbeatOk = await addSpocStudyHeartbeat();
    if (!heartbeatOk) {
      pushSpocLog("创建学习记录失败，将继续尝试上报进度。", "warning");
    }

    clearSpocSimulationTimer();
    const tickMs = Math.max(250, Math.floor(state.spoc.simInterval * 1000));
    state.spoc.timer = window.setInterval(async () => {
      if (!state.spoc.isRunning || state.spoc.simTickBusy) return;
      state.spoc.simTickBusy = true;

      try {
        const nextPlayTime = Math.min(
          state.spoc.simDuration,
          state.spoc.simPlayTime + state.spoc.simInterval * state.spoc.simSpeed
        );
        const nextProgress = Math.min(
          100,
          Math.floor((nextPlayTime / Math.max(1, state.spoc.simDuration)) * 100)
        );

        const ok = await updateSpocProgress(nextProgress, nextPlayTime);
        if (!ok) {
          stopSpocSimulation("进度上报失败，任务已停止。");
          return;
        }

        state.spoc.simPlayTime = nextPlayTime;
        state.spoc.simProgress = nextProgress;

        const remain = Math.max(0, (state.spoc.simDuration - nextPlayTime) / Math.max(0.01, state.spoc.simSpeed));
        const mins = Math.floor(remain / 60);
        const secs = Math.floor(remain % 60)
          .toString()
          .padStart(2, "0");
        state.spoc.simInfoText = `剩余约 ${mins}:${secs} | ${state.spoc.simSpeed}x | ${state.spoc.simInterval}s`;

        if (nextProgress >= 100) {
          clearSpocSimulationTimer();
          state.spoc.isRunning = false;
          state.spoc.simStatusText = "模拟上报完成。";
          state.spoc.simInfoText = "当前资源已标记为完成。";
          pushSpocLog("模拟任务完成。", "success");
        }
        renderSpocUi();
      } catch (error) {
        console.error(error);
        stopSpocSimulation("出现异常，任务已停止。");
      } finally {
        state.spoc.simTickBusy = false;
      }
    }, tickMs);
  }

  function initSpocHelper() {
    const shell = createPanelShell(
      "bsh-spoc-panel",
      "SPOC 播放辅助",
      "状态、倍速、自动下一节与视频链接工具"
    );
    shell.body.innerHTML = `
      <div class="bsh-section">
        <div class="bsh-badges">
          <span class="bsh-badge" data-role="video-badge">播放器未就绪</span>
          <span class="bsh-badge" data-role="auto-next-badge">自动下一节关闭</span>
          <span class="bsh-badge" data-role="automation-badge">自动上报未就绪</span>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card bsh-card--status" data-role="status">正在检测当前页面中的视频播放器。</div>
        <div class="bsh-card bsh-card--status" style="margin-top: 8px;" data-role="sim-info">
          请先点击一次资源“查看”，捕获参数和认证头后再开始自动上报。
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-progress">
          <div class="bsh-progress__fill" data-role="progress-fill" style="width: 0%">0%</div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card--grid">
          <div class="bsh-metric">
            <div class="bsh-metric__label">视频时长</div>
            <div class="bsh-metric__value" data-role="duration">待检测</div>
          </div>
          <div class="bsh-metric">
            <div class="bsh-metric__label">当前进度</div>
            <div class="bsh-metric__value" data-role="current-time">00:00:00</div>
          </div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card bsh-card--status">
          <div style="font-size: 11px; color: var(--bsh-text-muted);">当前资源</div>
          <div data-role="resource-title" style="margin-top: 6px; color: var(--bsh-text); word-break: break-word;">待识别</div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card">
          <div class="bsh-rate-presets" data-role="rate-presets"></div>
          <div class="bsh-field-row" style="margin-top: 10px;">
            <input type="number" class="bsh-input" min="0.25" max="16" step="0.25" value="1.5" data-role="rate-input">
            <button type="button" class="bsh-btn" data-action="apply-rate">应用倍速</button>
            <button type="button" class="bsh-btn bsh-btn--ghost" data-action="refresh">刷新检测</button>
          </div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card">
          <div style="font-size: 11px; color: var(--bsh-text-muted);">自动上报配置</div>
          <div class="bsh-field-row" style="margin-top: 10px;">
            <input type="number" class="bsh-input" min="1" max="43200" step="1" value="" style="width: 84px;" data-role="sim-duration">
            <span style="font-size: 11px; color: var(--bsh-text-muted);">时长(秒)</span>
          </div>
          <div class="bsh-field-row" style="margin-top: 8px;">
            <input type="number" class="bsh-input" min="0.25" max="100" step="0.25" value="10" style="width: 72px;" data-role="sim-speed">
            <span style="font-size: 11px; color: var(--bsh-text-muted);">速度(x)</span>
            <input type="number" class="bsh-input" min="0.5" max="30" step="0.5" value="1" style="width: 72px;" data-role="sim-interval">
            <span style="font-size: 11px; color: var(--bsh-text-muted);">心跳(s)</span>
          </div>
          <label style="display: flex; align-items: center; gap: 8px; margin-top: 10px; font-size: 11px; color: var(--bsh-text-muted);">
            <input type="checkbox" data-role="fast-mode">
            一键完成模式（高风险）
          </label>
          <div class="bsh-actions" style="margin-top: 10px;">
            <button type="button" class="bsh-btn bsh-btn--primary" data-action="start-sim">开始</button>
            <button type="button" class="bsh-btn bsh-btn--ghost" data-action="stop-sim" disabled>停止</button>
          </div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-actions">
          <button type="button" class="bsh-btn bsh-btn--warn" data-action="toggle-auto-next">开启自动下一节</button>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-card">
          <div class="bsh-link-tools">
            <button type="button" class="bsh-btn" data-action="open-link" disabled>打开视频链接</button>
            <button type="button" class="bsh-btn bsh-btn--ghost" data-action="copy-link" disabled>复制视频链接</button>
          </div>
          <div class="bsh-source" style="margin-top: 10px;" data-role="source">尚未读取到视频地址。</div>
        </div>
      </div>
      <div class="bsh-section">
        <div class="bsh-log" data-role="log"></div>
      </div>
    `;

    const ratePresetHost = shell.body.querySelector('[data-role="rate-presets"]');
    SPOC_RATE_PRESETS.forEach((rate) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "bsh-btn bsh-btn--ghost";
      button.dataset.rate = String(rate);
      button.textContent = `${rate}x`;
      button.addEventListener("click", () => applySpocRate(rate));
      ratePresetHost.appendChild(button);
    });

    state.spoc.ui = {
      shell,
      videoBadge: shell.body.querySelector('[data-role="video-badge"]'),
      autoNextBadge: shell.body.querySelector('[data-role="auto-next-badge"]'),
      automationBadge: shell.body.querySelector('[data-role="automation-badge"]'),
      status: shell.body.querySelector('[data-role="status"]'),
      simInfo: shell.body.querySelector('[data-role="sim-info"]'),
      progressFill: shell.body.querySelector('[data-role="progress-fill"]'),
      duration: shell.body.querySelector('[data-role="duration"]'),
      currentTime: shell.body.querySelector('[data-role="current-time"]'),
      resourceTitle: shell.body.querySelector('[data-role="resource-title"]'),
      rateInput: shell.body.querySelector('[data-role="rate-input"]'),
      simDurationInput: shell.body.querySelector('[data-role="sim-duration"]'),
      simSpeedInput: shell.body.querySelector('[data-role="sim-speed"]'),
      simIntervalInput: shell.body.querySelector('[data-role="sim-interval"]'),
      fastModeInput: shell.body.querySelector('[data-role="fast-mode"]'),
      openLinkButton: shell.body.querySelector('[data-action="open-link"]'),
      copyLinkButton: shell.body.querySelector('[data-action="copy-link"]'),
      toggleAutoNextButton: shell.body.querySelector('[data-action="toggle-auto-next"]'),
      refreshButton: shell.body.querySelector('[data-action="refresh"]'),
      applyRateButton: shell.body.querySelector('[data-action="apply-rate"]'),
      startSimButton: shell.body.querySelector('[data-action="start-sim"]'),
      stopSimButton: shell.body.querySelector('[data-action="stop-sim"]'),
      log: shell.body.querySelector('[data-role="log"]'),
      source: shell.body.querySelector('[data-role="source"]'),
      ratePresetButtons: Array.from(ratePresetHost.querySelectorAll("button")),
    };

    state.spoc.ui.openLinkButton.addEventListener("click", () => {
      if (state.spoc.currentSrc) {
        window.open(state.spoc.currentSrc, "_blank", "noopener,noreferrer");
      }
    });

    state.spoc.ui.copyLinkButton.addEventListener("click", async () => {
      if (!state.spoc.currentSrc) return;
      try {
        await navigator.clipboard.writeText(state.spoc.currentSrc);
        pushSpocLog("已复制当前视频链接。", "success");
      } catch (error) {
        pushSpocLog("复制链接失败，请检查浏览器剪贴板权限。", "warning");
      }
    });

    state.spoc.ui.toggleAutoNextButton.addEventListener("click", () => {
      state.spoc.autoNextEnabled = !state.spoc.autoNextEnabled;
      renderSpocUi();
      pushSpocLog(
        state.spoc.autoNextEnabled ? "已开启自动下一节。" : "已关闭自动下一节。",
        "info"
      );
    });

    state.spoc.ui.refreshButton.addEventListener("click", refreshSpocState);
    state.spoc.ui.applyRateButton.addEventListener("click", () => {
      const value = Number(state.spoc.ui.rateInput.value || state.spoc.preferredRate);
      applySpocRate(value);
    });

    state.spoc.ui.startSimButton.addEventListener("click", () => {
      startSpocSimulation();
    });

    state.spoc.ui.stopSimButton.addEventListener("click", () => {
      stopSpocSimulation("已手动停止任务。");
    });

    [state.spoc.ui.simDurationInput, state.spoc.ui.simSpeedInput, state.spoc.ui.simIntervalInput].forEach((input) => {
      input.addEventListener("change", () => {
        syncSpocAutomationConfigFromUi();
        renderSpocUi();
      });
    });

    state.spoc.ui.fastModeInput.addEventListener("change", () => {
      syncSpocAutomationConfigFromUi();
      renderSpocUi();
      if (state.spoc.fastMode) {
        pushSpocLog("已开启一键完成模式。", "warning");
      } else {
        pushSpocLog("已关闭一键完成模式。", "info");
      }
    });

    syncSpocAutomationConfigFromUi();
    installSpocRequestInterceptor();

    pushSpocLog("播放器辅助已加载，正在检测当前页面视频。", "success");
    refreshSpocState();
    startSpocRateEnforcer();
    state.spoc.refreshTimer = window.setInterval(refreshSpocState, 1200);
  }

  async function refreshSpocState() {
    const resource = getCurrentSpocResource();
    state.spoc.currentResourceTitle = resource?.title || "";

    const video = findBestVideoElement();
    state.spoc.currentVideo = video;
    bindSpocVideo(video);

    const duration = await detectVideoDuration(video);
    state.spoc.currentDuration = duration;
    state.spoc.currentSrc = getVideoSource(video);

    renderSpocUi();
  }

  function renderSpocUi() {
    const ui = state.spoc.ui;
    if (!ui) return;

    const video = state.spoc.currentVideo;
    const realDuration = state.spoc.currentDuration;
    const realCurrentTime = video?.currentTime || 0;
    const realProgress = realDuration > 0 ? Math.min(100, Math.floor((realCurrentTime / realDuration) * 100)) : 0;

    const running = state.spoc.isRunning;
    const displayDuration = running ? Math.max(state.spoc.simDuration || 0, realDuration || 0) : realDuration;
    const displayCurrentTime = running ? state.spoc.simPlayTime : realCurrentTime;
    const displayProgress = running ? state.spoc.simProgress : realProgress;

    const hasParams = Boolean(state.spoc.params?.kcid && state.spoc.params?.kcnrid && state.spoc.params?.ssmlid);
    const hasToken = Boolean(getSpocHeaderValue("Token"));
    const automationReady = hasParams && hasToken;

    updateBadge(ui.videoBadge, Boolean(video), video ? "播放器已就绪" : "播放器未就绪");
    updateBadge(
      ui.autoNextBadge,
      state.spoc.autoNextEnabled,
      state.spoc.autoNextEnabled ? "自动下一节已开启" : "自动下一节关闭",
      !state.spoc.autoNextEnabled
    );

    if (ui.automationBadge) {
      if (running) {
        updateBadge(ui.automationBadge, true, "自动上报运行中");
      } else if (automationReady) {
        updateBadge(ui.automationBadge, true, "自动上报已就绪");
      } else {
        updateBadge(ui.automationBadge, false, "自动上报未就绪", true);
      }
    }

    if (running) {
      ui.status.textContent = state.spoc.simStatusText || "正在模拟上报进度...";
    } else {
      ui.status.textContent = video
        ? "已检测到视频播放器，可使用倍速、自动下一节和链接工具。"
        : "当前页面尚未检测到可用播放器，请打开具体视频后再试。";
    }

    if (ui.simInfo) {
      ui.simInfo.textContent =
        state.spoc.simInfoText ||
        (automationReady ? "参数和认证头已捕获，可开始自动上报。" : "请点击一次“查看”以捕获参数和认证头。");
    }

    ui.progressFill.style.width = `${Math.max(0, Math.min(100, displayProgress))}%`;
    ui.progressFill.textContent = `${Math.max(0, Math.min(100, displayProgress))}%`;
    ui.duration.textContent = displayDuration > 0 ? formatDurationLabel(displayDuration) : "待检测";
    ui.currentTime.textContent = formatClockTime(displayCurrentTime);
    ui.resourceTitle.textContent = state.spoc.currentResourceTitle || "待识别";
    ui.source.textContent = state.spoc.currentSrc || "尚未读取到视频地址。";
    ui.openLinkButton.disabled = !state.spoc.currentSrc;
    ui.copyLinkButton.disabled = !state.spoc.currentSrc;
    ui.toggleAutoNextButton.textContent = state.spoc.autoNextEnabled ? "关闭自动下一节" : "开启自动下一节";
    ui.startSimButton.disabled = running || !automationReady;
    ui.stopSimButton.disabled = !running;

    ui.ratePresetButtons.forEach((button) => {
      const rate = Number(button.dataset.rate || 1);
      const active = Math.abs(rate - state.spoc.preferredRate) < 0.001;
      button.classList.toggle("bsh-btn--active", active);
    });

    ui.rateInput.value = String(state.spoc.preferredRate);
    ui.simSpeedInput.value = String(state.spoc.simSpeed);
    ui.simIntervalInput.value = String(state.spoc.simInterval);
    if (state.spoc.simDuration > 0) {
      ui.simDurationInput.value = String(Math.floor(state.spoc.simDuration));
    }
    ui.fastModeInput.checked = Boolean(state.spoc.fastMode);
    renderSpocLog();
  }

  function getCurrentSpocResource() {
    const activeTitle = document.querySelector('.title-active .panel-title, .title-active[title], .title-active');
    const title = activeTitle?.getAttribute?.('title') || activeTitle?.textContent?.trim() || "";
    return title ? { title } : null;
  }

  function findCurrentSpocResourcePanel() {
    return document.querySelector('.title-active')?.closest('.collapse-panel') || null;
  }

  function findCurrentSpocViewButton() {
    return Array.from(document.querySelectorAll('.course-r-module button.ivu-btn.ivu-btn-info.ivu-btn-small')).find((button) => isVisible(button)) || null;
  }

  async function openNextSpocResource() {
    const currentPanel = findCurrentSpocResourcePanel();
    if (!currentPanel || !currentPanel.parentElement) {
      pushSpocLog("未识别到当前资源节点，无法切换下一项。", "warning");
      return false;
    }

    const siblings = Array.from(currentPanel.parentElement.children).filter((el) => el.classList?.contains('collapse-panel'));
    const currentIndex = siblings.indexOf(currentPanel);
    if (currentIndex < 0 || currentIndex >= siblings.length - 1) {
      pushSpocLog("已经是当前分组的最后一个资源。", "warning");
      return false;
    }

    const nextPanel = siblings[currentIndex + 1];
    const nextTitle = nextPanel.querySelector('.panel-title')?.textContent?.trim() || nextPanel.textContent?.trim() || '下一项';
    const nextHeader = nextPanel.querySelector('.ivu-collapse-header');
    if (nextHeader) {
      nextHeader.click();
      await sleep(500);
    }

    let switched = false;
    for (let attempt = 0; attempt < 8; attempt += 1) {
      const currentTitle = getCurrentSpocResource()?.title || '';
      if (currentTitle && currentTitle.includes(nextTitle)) {
        switched = true;
        break;
      }
      await sleep(250);
    }

    if (!switched) {
      pushSpocLog(`未能切换到下一资源：${nextTitle}`, "warning");
      return false;
    }

    const nextViewButton = findCurrentSpocViewButton();
    if (!nextViewButton) {
      pushSpocLog("已切换到下一资源，但未找到右侧对应的“查看”按钮。", "warning");
      return false;
    }

    nextViewButton.click();
    pushSpocLog(`已尝试打开下一资源：${nextTitle}`, "success");
    return true;
  }

  function renderSpocLog() {
    const ui = state.spoc.ui;
    if (!ui) return;
    ui.log.innerHTML = state.spoc.logs
      .map(
        (entry) =>
          `<div class="bsh-log__item ${entry.tone ? `is-${entry.tone}` : ""}">[${escapeHtml(
            entry.time
          )}] ${escapeHtml(entry.message)}</div>`
      )
      .join("");
  }

  function pushSpocLog(message, tone = "info") {
    state.spoc.logs.unshift({
      message,
      tone,
      time: new Date().toLocaleTimeString("zh-CN"),
    });
    state.spoc.logs = state.spoc.logs.slice(0, 8);
    renderSpocLog();
  }

  function findBestVideoElement() {
    const modalVideo = document.querySelector('.ivu-modal-wrap:not(.ivu-modal-hidden) video, .ivu-modal video, #myVideo_html5_api');
    if (modalVideo) return modalVideo;

    const videos = Array.from(document.querySelectorAll("video"));
    if (!videos.length) return null;
    return (
      videos.find((video) => Boolean(getVideoSource(video)) && isVisible(video)) ||
      videos.find((video) => isVisible(video)) ||
      videos[0]
    );
  }

  function isVisible(element) {
    if (!(element instanceof HTMLElement)) return false;
    const style = window.getComputedStyle(element);
    return style.display !== "none" && style.visibility !== "hidden" && element.offsetWidth > 0 && element.offsetHeight > 0;
  }

  function getVideoSource(video) {
    if (!video) return "";
    return video.currentSrc || video.src || video.querySelector("source")?.src || "";
  }

  async function detectVideoDuration(video) {
    if (video?.duration && Number.isFinite(video.duration) && video.duration > 0) {
      return Math.floor(video.duration);
    }

    const durationElement = document.querySelector(
      ".vjs-duration-display, .duration, [class*='duration'], .video-time"
    );
    if (durationElement?.textContent) {
      const match = durationElement.textContent.match(/(\d+):(\d+):?(\d+)?/);
      if (match) {
        if (match[3]) {
          return Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3]);
        }
        return Number(match[1]) * 60 + Number(match[2]);
      }
    }

    const source = getVideoSource(video);
    if (source && source.includes(".mp4")) {
      return new Promise((resolve) => {
        const probe = document.createElement("video");
        let resolved = false;
        const finish = (value) => {
          if (resolved) return;
          resolved = true;
          resolve(value);
        };
        probe.preload = "metadata";
        probe.onloadedmetadata = () => finish(Math.floor(probe.duration || 0));
        probe.onerror = () => finish(0);
        probe.src = source;
        setTimeout(() => finish(0), 5000);
      });
    }

    return 0;
  }

  function bindSpocVideo(video) {
    if (!video || video.dataset.bshBound === "true") {
      if (video) {
        state.spoc.currentVideo = video;
        ensureSpocPlaybackRate(true);
      }
      return;
    }
    video.dataset.bshBound = "true";

    video.addEventListener("loadedmetadata", async () => {
      state.spoc.currentVideo = video;
      state.spoc.currentDuration = await detectVideoDuration(video);
      ensureSpocPlaybackRate(true);
      renderSpocUi();
    });

    video.addEventListener("timeupdate", renderSpocUi);
    video.addEventListener("ratechange", () => {
      ensureSpocPlaybackRate(true);
      renderSpocUi();
    });

    video.addEventListener("play", () => {
      ensureSpocPlaybackRate(true);
      renderSpocUi();
    });

    video.addEventListener("ended", async () => {
      pushSpocLog("当前视频播放结束。", "info");
      if (state.spoc.autoNextEnabled) {
        await tryPlayNextVideo();
      }
    });

    state.spoc.currentVideo = video;
    ensureSpocPlaybackRate(true);
  }

  function applySpocRate(rate) {
    const normalized = Number(rate);
    if (!Number.isFinite(normalized) || normalized < 0.25 || normalized > 16) {
      pushSpocLog("请输入 0.25 到 16 之间的倍速。", "warning");
      return;
    }
    state.spoc.preferredRate = normalized;
    ensureSpocPlaybackRate(true);
    pushSpocLog(`已设置并强制保持播放倍速为 ${normalized}x。`, "success");
    renderSpocUi();
  }

  function findOpenSpocModal() {
    const modals = Array.from(document.querySelectorAll('.ivu-modal-wrap:not(.ivu-modal-hidden)'));
    return modals.find((modal) => {
      const title = modal.querySelector('.ivu-modal-header-inner')?.textContent?.trim() || '';
      return title.startsWith('查看【');
    }) || null;
  }

  function closeOpenSpocModal() {
    const modal = findOpenSpocModal();
    if (!modal) return false;
    const closeButton = modal.querySelector('.ivu-modal-footer .ivu-btn, .vjs-close-button');
    if (!closeButton) return false;
    closeButton.click();
    return true;
  }

  async function tryPlayNextVideo() {
    const closed = closeOpenSpocModal();
    if (closed) {
      await sleep(400);
    }

    const opened = await openNextSpocResource();
    if (!opened) {
      return;
    }

    await sleep(1200);
    const nextVideo = findBestVideoElement();
    if (nextVideo) {
      state.spoc.currentVideo = nextVideo;
      bindSpocVideo(nextVideo);
      if (state.spoc.preferredRate && Number.isFinite(state.spoc.preferredRate)) {
        nextVideo.playbackRate = state.spoc.preferredRate;
      }
      if (typeof nextVideo.play === 'function') {
        nextVideo.play().catch(() => {});
      }
    }

    refreshSpocState();
  }

  function generateUuid() {
    const chunk = () => (((1 + Math.random()) * 0x10000) | 0).toString(16).slice(1);
    return `${chunk()}${chunk()}-${chunk()}-${chunk()}-${chunk()}-${chunk()}${chunk()}${chunk()}`;
  }

  function md5Hex(text) {
    let hexcase = 0;
    let chrsz = 8;

    const safeAdd = (x, y) => {
      const lsw = (x & 0xffff) + (y & 0xffff);
      const msw = (x >> 16) + (y >> 16) + (lsw >> 16);
      return (msw << 16) | (lsw & 0xffff);
    };

    const bitRol = (num, cnt) => (num << cnt) | (num >>> (32 - cnt));
    const md5Cmn = (q, a, b, x, s, t) => safeAdd(bitRol(safeAdd(safeAdd(a, q), safeAdd(x, t)), s), b);
    const md5Ff = (a, b, c, d, x, s, t) => md5Cmn((b & c) | (~b & d), a, b, x, s, t);
    const md5Gg = (a, b, c, d, x, s, t) => md5Cmn((b & d) | (c & ~d), a, b, x, s, t);
    const md5Hh = (a, b, c, d, x, s, t) => md5Cmn(b ^ c ^ d, a, b, x, s, t);
    const md5Ii = (a, b, c, d, x, s, t) => md5Cmn(c ^ (b | ~d), a, b, x, s, t);

    const str2binl = (str) => {
      const bin = [];
      const mask = (1 << chrsz) - 1;
      for (let index = 0; index < str.length * chrsz; index += chrsz) {
        bin[index >> 5] |= (str.charCodeAt(index / chrsz) & mask) << (index % 32);
      }
      return bin;
    };

    const binl2hex = (binarray) => {
      const hexTab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
      let output = "";
      for (let index = 0; index < binarray.length * 4; index += 1) {
        output +=
          hexTab.charAt((binarray[index >> 2] >> (((index % 4) * 8) + 4)) & 0x0f) +
          hexTab.charAt((binarray[index >> 2] >> ((index % 4) * 8)) & 0x0f);
      }
      return output;
    };

    const coreMd5 = (x, len) => {
      x[len >> 5] |= 0x80 << (len % 32);
      x[(((len + 64) >>> 9) << 4) + 14] = len;
      let a = 1732584193;
      let b = -271733879;
      let c = -1732584194;
      let d = 271733878;

      for (let index = 0; index < x.length; index += 16) {
        const oldA = a;
        const oldB = b;
        const oldC = c;
        const oldD = d;

        a = md5Ff(a, b, c, d, x[index + 0], 7, -680876936);
        d = md5Ff(d, a, b, c, x[index + 1], 12, -389564586);
        c = md5Ff(c, d, a, b, x[index + 2], 17, 606105819);
        b = md5Ff(b, c, d, a, x[index + 3], 22, -1044525330);
        a = md5Ff(a, b, c, d, x[index + 4], 7, -176418897);
        d = md5Ff(d, a, b, c, x[index + 5], 12, 1200080426);
        c = md5Ff(c, d, a, b, x[index + 6], 17, -1473231341);
        b = md5Ff(b, c, d, a, x[index + 7], 22, -45705983);
        a = md5Ff(a, b, c, d, x[index + 8], 7, 1770035416);
        d = md5Ff(d, a, b, c, x[index + 9], 12, -1958414417);
        c = md5Ff(c, d, a, b, x[index + 10], 17, -42063);
        b = md5Ff(b, c, d, a, x[index + 11], 22, -1990404162);
        a = md5Ff(a, b, c, d, x[index + 12], 7, 1804603682);
        d = md5Ff(d, a, b, c, x[index + 13], 12, -40341101);
        c = md5Ff(c, d, a, b, x[index + 14], 17, -1502002290);
        b = md5Ff(b, c, d, a, x[index + 15], 22, 1236535329);
        a = md5Gg(a, b, c, d, x[index + 1], 5, -165796510);
        d = md5Gg(d, a, b, c, x[index + 6], 9, -1069501632);
        c = md5Gg(c, d, a, b, x[index + 11], 14, 643717713);
        b = md5Gg(b, c, d, a, x[index + 0], 20, -373897302);
        a = md5Gg(a, b, c, d, x[index + 5], 5, -701558691);
        d = md5Gg(d, a, b, c, x[index + 10], 9, 38016083);
        c = md5Gg(c, d, a, b, x[index + 15], 14, -660478335);
        b = md5Gg(b, c, d, a, x[index + 4], 20, -405537848);
        a = md5Gg(a, b, c, d, x[index + 9], 5, 568446438);
        d = md5Gg(d, a, b, c, x[index + 14], 9, -1019803690);
        c = md5Gg(c, d, a, b, x[index + 3], 14, -187363961);
        b = md5Gg(b, c, d, a, x[index + 8], 20, 1163531501);
        a = md5Gg(a, b, c, d, x[index + 13], 5, -1444681467);
        d = md5Gg(d, a, b, c, x[index + 2], 9, -51403784);
        c = md5Gg(c, d, a, b, x[index + 7], 14, 1735328473);
        b = md5Gg(b, c, d, a, x[index + 12], 20, -1926607734);
        a = md5Hh(a, b, c, d, x[index + 5], 4, -378558);
        d = md5Hh(d, a, b, c, x[index + 8], 11, -2022574463);
        c = md5Hh(c, d, a, b, x[index + 11], 16, 1839030562);
        b = md5Hh(b, c, d, a, x[index + 14], 23, -35309556);
        a = md5Hh(a, b, c, d, x[index + 1], 4, -1530992060);
        d = md5Hh(d, a, b, c, x[index + 4], 11, 1272893353);
        c = md5Hh(c, d, a, b, x[index + 7], 16, -155497632);
        b = md5Hh(b, c, d, a, x[index + 10], 23, -1094730640);
        a = md5Hh(a, b, c, d, x[index + 13], 4, 681279174);
        d = md5Hh(d, a, b, c, x[index + 0], 11, -358537222);
        c = md5Hh(c, d, a, b, x[index + 3], 16, -722521979);
        b = md5Hh(b, c, d, a, x[index + 6], 23, 76029189);
        a = md5Hh(a, b, c, d, x[index + 9], 4, -640364487);
        d = md5Hh(d, a, b, c, x[index + 12], 11, -421815835);
        c = md5Hh(c, d, a, b, x[index + 15], 16, 530742520);
        b = md5Hh(b, c, d, a, x[index + 2], 23, -995338651);
        a = md5Ii(a, b, c, d, x[index + 0], 6, -198630844);
        d = md5Ii(d, a, b, c, x[index + 7], 10, 1126891415);
        c = md5Ii(c, d, a, b, x[index + 14], 15, -1416354905);
        b = md5Ii(b, c, d, a, x[index + 5], 21, -57434055);
        a = md5Ii(a, b, c, d, x[index + 12], 6, 1700485571);
        d = md5Ii(d, a, b, c, x[index + 3], 10, -1894986606);
        c = md5Ii(c, d, a, b, x[index + 10], 15, -1051523);
        b = md5Ii(b, c, d, a, x[index + 1], 21, -2054922799);
        a = md5Ii(a, b, c, d, x[index + 8], 6, 1873313359);
        d = md5Ii(d, a, b, c, x[index + 15], 10, -30611744);
        c = md5Ii(c, d, a, b, x[index + 6], 15, -1560198380);
        b = md5Ii(b, c, d, a, x[index + 13], 21, 1309151649);
        a = md5Ii(a, b, c, d, x[index + 4], 6, -145523070);
        d = md5Ii(d, a, b, c, x[index + 11], 10, -1120210379);
        c = md5Ii(c, d, a, b, x[index + 2], 15, 718787259);
        b = md5Ii(b, c, d, a, x[index + 9], 21, -343485551);

        a = safeAdd(a, oldA);
        b = safeAdd(b, oldB);
        c = safeAdd(c, oldC);
        d = safeAdd(d, oldD);
      }
      return [a, b, c, d];
    };

    return binl2hex(coreMd5(str2binl(text), text.length * chrsz));
  }
})();