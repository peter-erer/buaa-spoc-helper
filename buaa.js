// ==UserScript==
// @name         智学北航学习助手
// @namespace    http://tampermonkey.net/
// @version      3.1
// @description  一键导出智学北航课程字幕(SRT/TXT)及PPT课件(PDF)，精准适配search-ppt接口
// @author       Peter Sheild
// @match        *://*.classroom.msa.buaa.edu.cn/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // 存储数据
    let subtitleData = null;
    let pptData = null;

    // --- 1. 核心拦截逻辑 ---
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function(method, url) {
        this._url = url;
        return originalOpen.apply(this, arguments);
    };

    XMLHttpRequest.prototype.send = function() {
        this.addEventListener('load', function() {
            if (!this._url) return;

            // 字幕通道
            if (this._url.includes('search-trans-result')) {
                try {
                    const response = JSON.parse(this.responseText);
                    if (response && response.list) {
                        subtitleData = response.list;
                        updateBtnStatus('subtitle', true);
                        console.log("✅ 字幕数据已捕获");
                    }
                } catch (e) {}
            }
            
            // PPT通道
            if (this._url.includes('search-ppt')) {
                try {
                    const response = JSON.parse(this.responseText);
                    if (response && response.list) {
                        pptData = response.list;
                        updateBtnStatus('ppt', true);
                        console.log("✅ PPT 数据已捕获 (条数: " + pptData.length + ")");
                    }
                } catch (e) {
                    console.error("PPT解析失败", e);
                }
            }
        });
        return originalSend.apply(this, arguments);
    };

    // --- 2. UI 创建 ---
    function initUI() {
        let container = document.getElementById('buaa-export-container');
        if (!container) {
            container = document.createElement('div');
            container.id = 'buaa-export-container';
            Object.assign(container.style, {
                position: 'fixed', top: '150px', right: '50px', zIndex: '9999',
                display: 'flex', flexDirection: 'column', gap: '8px',
                padding: '12px', background: 'rgba(0, 0, 0, 0.75)', borderRadius: '8px',
                color: 'white', fontSize: '12px', backdropFilter: 'blur(5px)',
                boxShadow: '0 4px 15px rgba(0,0,0,0.3)', cursor: 'move', userSelect: 'none',
                width: '140px'
            });
            
            const title = document.createElement('div');
            title.innerText = '::: 学习助手 :::';
            title.style.cssText = 'text-align:center; margin-bottom:5px; color:#aaa; font-weight:bold;';
            container.appendChild(title);

            const btnGroup = document.createElement('div');
            btnGroup.id = 'buaa-btn-group';
            container.appendChild(btnGroup);

            document.body.appendChild(container);
            makeDraggable(container);
            
            renderButtons();
        }
    }

    function renderButtons() {
        const group = document.getElementById('buaa-btn-group');
        group.innerHTML = '';
        createBtn(group, 'subtitle', '🎬 导出 SRT 字幕', () => processSubtitle('srt'));
        createBtn(group, 'subtitle', '📝 导出 TXT 笔记', () => processSubtitle('txt'));
        createBtn(group, 'ppt',      '📊 导出 PPT (PDF)', () => processPPT());
    }

    function createBtn(parent, type, text, onClick) {
        const btn = document.createElement('button');
        btn.id = `btn-${type}`;
        btn.innerText = text;
        Object.assign(btn.style, {
            display: 'block', width: '100%', padding: '8px 10px',
            border: 'none', borderRadius: '4px',
            background: '#555', color: '#aaa', cursor: 'not-allowed',
            transition: 'all 0.3s'
        });
        
        if ((type === 'subtitle' && subtitleData) || (type === 'ppt' && pptData)) {
            activateBtnStyle(btn);
        }

        btn.onclick = () => {
            if ((type === 'subtitle' && subtitleData) || (type === 'ppt' && pptData)) {
                onClick();
            } else {
                alert(`等待数据加载中...\n\n请尝试点击一下页面左侧的"${type === 'ppt' ? 'PPT' : '语音'}"标签页以触发数据请求。`);
            }
        };
        btn.onmousedown = (e) => e.stopPropagation(); 
        parent.appendChild(btn);
    }

    function updateBtnStatus(type, ready) {
        if (!ready) return;
        const btns = document.querySelectorAll(`button[id^='btn-${type}']`);
        btns.forEach(btn => activateBtnStyle(btn));
    }

    function activateBtnStyle(btn) {
        btn.style.background = '#005596';
        btn.style.color = 'white';
        btn.style.cursor = 'pointer';
    }

    // --- 3. 字幕处理 (TXT合并/SRT不合并) ---
    function processSubtitle(type) {
        if (!subtitleData) return;
        
        let rawLines = subtitleData.reduce((acc, chap) => acc.concat(chap.all_content || []), []);
        let linesToProcess = rawLines;

        // 仅在 TXT 模式下合并短句
        if (type === 'txt') {
            let merged = [];
            if (rawLines.length > 0) {
                let cur = { ...rawLines[0] };
                for (let i = 1; i < rawLines.length; i++) {
                    let nxt = rawLines[i];
                    let gap = nxt.BeginSec - (cur.EndSec || cur.BeginSec);
                    // 间隔小于1秒且字数少于20才合并
                    if (gap < 1.0 && cur.Text.length < 20) {
                        cur.Text += "，" + nxt.Text; 
                        cur.EndSec = nxt.EndSec;
                    } else { 
                        merged.push(cur); 
                        cur = { ...nxt }; 
                    }
                }
                merged.push(cur);
                linesToProcess = merged;
            }
        }

        let content = type === 'txt' ? `课程笔记 - ${document.title}\n\n` : "";
        linesToProcess.forEach((line, i) => {
            if (type === 'srt') {
                content += `${i+1}\n${formatTime(line.BeginSec)} --> ${formatTime(line.EndSec||line.BeginSec+2)}\n${line.Text}\n\n`;
            } else {
                content += `[${formatTime(line.BeginSec).substr(0,8)}] ${line.Text}\n`;
            }
        });
        downloadFile(content, type);
    }

    // --- 4. PPT 处理 (已修正解析逻辑) ---
    function processPPT() {
        if (!pptData || pptData.length === 0) {
            alert("未找到 PPT 图片数据！");
            return;
        }

        const printWindow = window.open('', '_blank');
        if (!printWindow) {
            alert("请允许本网站弹出窗口，否则无法生成 PDF！");
            return;
        }

        let htmlContent = `
            <html>
            <head>
                <title>PPT导出 - ${document.title}</title>
                <style>
                    body { font-family: sans-serif; background: #f0f0f0; margin: 0; padding: 20px; }
                    .slide-container { 
                        max-width: 1000px; margin: 0 auto; background: white; 
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px;
                        page-break-inside: avoid;
                    }
                    img { width: 100%; display: block; }
                    .meta { padding: 10px; color: #666; font-size: 14px; border-top: 1px solid #eee; }
                    @media print {
                        body { background: white; padding: 0; }
                        .slide-container { box-shadow: none; margin-bottom: 0; page-break-after: always; }
                        @page { margin: 0; }
                    }
                </style>
            </head>
            <body>
                <div style="text-align:center; margin-bottom:20px; padding:20px;">
                    <h1>${document.title} - 课程 PPT</h1>
                    <p style="color:red; font-weight:bold;">⚠️ 请在打印界面选择“另存为 PDF”</p>
                </div>
        `;

        pptData.forEach((slide, index) => {
            let imgUrl = "";
            // === 修正后的解析逻辑 ===
            try {
                // 解析 content 字符串
                if (slide.content) {
                    const contentObj = JSON.parse(slide.content);
                    imgUrl = contentObj.pptimgurl;
                }
            } catch (e) { console.error(e); }

            // 使用 created_sec 作为时间戳
            let timeStr = formatTime(slide.created_sec || slide.BeginSec || 0).substr(0, 8);

            if (imgUrl) {
                htmlContent += `
                    <div class="slide-container">
                        <img src="${imgUrl}" loading="lazy">
                        <div class="meta">第 ${index + 1} 页 | 视频时间: ${timeStr}</div>
                    </div>
                `;
            }
        });

        htmlContent += `
            <script>
                window.onload = function() {
                    // 图片加载缓冲
                    setTimeout(() => { window.print(); }, 1500);
                };
            <\/script>
            </body></html>
        `;

        printWindow.document.write(htmlContent);
        printWindow.document.close();
    }

    // --- 5. 辅助工具 ---
    function formatTime(seconds) {
        if (!seconds) return "00:00:00";
        const date = new Date(0);
        date.setSeconds(seconds);
        return date.toISOString().substr(11, 8) + ",000";
    }

    function downloadFile(content, ext) {
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const safeTitle = document.title.replace(/[\\/:*?"<>|]/g, "_").substr(0, 30);
        a.href = url;
        a.download = `${safeTitle}_${ext === 'srt' ? '字幕' : '笔记'}.${ext}`;
        a.click();
        URL.revokeObjectURL(url);
    }

    function makeDraggable(el) {
        let isDragging = false;
        let startX, startY, initialLeft, initialTop;
        el.addEventListener('mousedown', (e) => {
            isDragging = true;
            startX = e.clientX; startY = e.clientY;
            initialLeft = el.offsetLeft; initialTop = el.offsetTop;
            el.style.cursor = 'grabbing';
        });
        document.addEventListener('mousemove', (e) => {
            if (isDragging) {
                el.style.left = `${initialLeft + (e.clientX - startX)}px`;
                el.style.top = `${initialTop + (e.clientY - startY)}px`;
                el.style.right = 'auto';
            }
        });
        document.addEventListener('mouseup', () => { isDragging = false; el.style.cursor = 'move'; });
    }

    window.addEventListener('load', initUI);
})();