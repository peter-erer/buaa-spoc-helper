// ==UserScript==
// @name         智学北航学习助手
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  一键导出PPT/语音识别字幕SRT或时间戳笔记TXT
// @author       Peter Sheild
// @match        *://*.classroom.msa.buaa.edu.cn/livingroom*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // 存储数据
    let subtitleData = null;
    let pptData = null;

    // --- 1. 核心拦截逻辑 (保持稳定) ---
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function(method, url) {
        this._url = url;
        return originalOpen.apply(this, arguments);
    };

    XMLHttpRequest.prototype.send = function() {
        this.addEventListener('load', function() {
            if (!this._url) return;

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

            if (this._url.includes('search-ppt')) {
                try {
                    const response = JSON.parse(this.responseText);
                    if (response && response.list) {
                        pptData = response.list;
                        updateBtnStatus('ppt', true);
                        console.log("✅ PPT 数据已捕获");
                    }
                } catch (e) { console.error(e); }
            }
        });
        return originalSend.apply(this, arguments);
    };

    // --- 2. UI 创建 (保持稳定) ---
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
        createBtn(group, 'ppt',      '📊 导出 PPT 讲义', () => processPPT());
        createBtn(group, 'subtitle', '🎬 导出 SRT 字幕', () => processSubtitle('srt'));
        createBtn(group, 'subtitle', '📝 导出 TXT 笔记', () => processSubtitle('txt'));
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

    // --- 3. 字幕处理 (保持稳定) ---
    function processSubtitle(type) {
        if (!subtitleData) return;

        let rawLines = subtitleData.reduce((acc, chap) => acc.concat(chap.all_content || []), []);
        let linesToProcess = rawLines;

        if (type === 'txt') {
            let merged = [];
            if (rawLines.length > 0) {
                let cur = { ...rawLines[0] };
                for (let i = 1; i < rawLines.length; i++) {
                    let nxt = rawLines[i];
                    let gap = nxt.BeginSec - (cur.EndSec || cur.BeginSec);
                    if (gap < 1.0 && cur.Text.length < 20) {
                        cur.Text += "，" + nxt.Text; cur.EndSec = nxt.EndSec;
                    } else { merged.push(cur); cur = { ...nxt }; }
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

    // --- 4. PPT 处理 (100%+不强制横向+border-box优雅布局+预加载) ---
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
                    * { box-sizing: border-box; }
                    body { font-family: "Microsoft YaHei", sans-serif; background: #f0f0f0; margin: 0; padding: 20px; }

                    /* 加载遮罩层样式 */
                    #loading-mask {
                        position: fixed; top: 0; left: 0; width: 100%; height: 100%;
                        background: rgba(255, 255, 255, 0.95); z-index: 9999;
                        display: flex; flex-direction: column;
                        justify-content: center; align-items: center;
                    }
                    #loading-text { font-size: 24px; color: #005596; font-weight: bold; margin-bottom: 10px;}
                    #loading-sub { font-size: 14px; color: #666; }

                    /* 文档样式 */
                    .doc-header {
                        text-align: center; margin-bottom: 20px; padding: 15px;
                        background: white; border-radius: 8px;
                    }
                    .slide-container {
                        width: 90%; max-width: 1000px; margin: 0 auto 30px auto;
                        background: white; border-radius: 4px; overflow: hidden;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                        border: 1px solid #ddd;
                    }
                    img { width: 100%; display: block; border-bottom: 1px solid #eee; }
                    .meta { padding: 10px 20px; color: #555; font-size: 13px; background: #fafafa; display: flex; justify-content: space-between; }

                    @media print {
                        @page { margin: 1cm; }
                        body { background: white; padding: 0; margin: 0; }
                        #loading-mask { display: none !important; } /* 打印时隐藏遮罩 */

                        .doc-header {
                            margin: 0 0 15px 0; padding: 0 0 10px 0;
                            border-bottom: 2px solid #333;
                            text-align: left; border-radius: 0;
                        }
                        .doc-header h1 { font-size: 18pt; margin: 0; }
                        .doc-header p { font-size: 10pt; margin: 5px 0 0 0; color: #666; }
                        .no-print { display: none !important; }

                        .slide-container {
                            width: 100%; max-width: 100%; margin: 0;
                            box-shadow: none; page-break-after: always; page-break-inside: avoid;
                            border: 1px solid #999;
                        }
                        img { max-width: 100%; max-height: 82vh; width: auto; margin: 0 auto; }
                        .meta {
                            border-top: 1px solid #999; padding: 8px 15px; font-size: 10pt;
                            background-color: #fafafa !important;
                            -webkit-print-color-adjust: exact; print-color-adjust: exact;
                        }
                    }
                </style>
            </head>
            <body>
                <div id="loading-mask">
                    <div id="loading-text">正在准备资源...</div>
                    <div id="loading-sub">请稍候，正在预加载高清图片以防止打印空白</div>
                </div>

                <div class="doc-header">
                    <h1>${document.title}</h1>
                    <p>课程讲义 | 生成时间: ${new Date().toLocaleDateString()}</p>
                    <p class="no-print" style="color:red; font-size:12px; margin-top:5px;">
                        💡 提示：推荐打印设置中手动选择 <b>“横向”</b>。
                    </p>
                </div>
        `;

        pptData.forEach((slide, index) => {
            let imgUrl = "";
            try {
                if (slide.content) {
                    const contentObj = JSON.parse(slide.content);
                    imgUrl = contentObj.pptimgurl;
                }
            } catch (e) {}
            let timeStr = formatTime(slide.created_sec || slide.BeginSec || 0).substr(0, 8);

            if (imgUrl) {
                htmlContent += `
                    <div class="slide-container">
                        <img src="${imgUrl}" class="ppt-img">
                        <div class="meta">
                            <span>第 ${index + 1} 页</span>
                            <span>⏱️ 视频时间点: ${timeStr}</span>
                        </div>
                    </div>
                `;
            }
        });

        htmlContent += `
            <script>
                window.onload = function() {
                    const images = document.querySelectorAll('.ppt-img');
                    const total = images.length;
                    let loaded = 0;
                    const textEl = document.getElementById('loading-text');

                    // 进度更新函数
                    function checkProgress() {
                        loaded++;
                        textEl.innerText = '正在加载图片 (' + loaded + '/' + total + ')...';

                        // 全部加载完毕
                        if (loaded >= total) {
                            textEl.innerText = '加载完成，正在唤起打印...';
                            setTimeout(() => {
                                document.getElementById('loading-mask').style.display = 'none';
                                window.print();
                            }, 500);
                        }
                    }

                    // 遍历所有图片，监听加载状态
                    if (total === 0) {
                        checkProgress(); // 只有文字的情况
                    } else {
                        images.forEach(img => {
                            if (img.complete) {
                                checkProgress(); // 已经缓存过的
                            } else {
                                img.onload = checkProgress;
                                img.onerror = checkProgress; // 即使图片挂了也继续，防止卡死
                            }
                        });
                    }
                };
            <\/script>
            </body></html>
        `;

        printWindow.document.write(htmlContent);
        printWindow.document.close();
    }
    // --- 5. 辅助工具 (保持稳定) ---
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