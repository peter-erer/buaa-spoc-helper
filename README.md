# 🎓 智学北航学习助手 (Smart Beihang Learning Assistant)

![Version](https://img.shields.io/badge/version-4.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)
[![Greasy Fork Downloads](https://img.shields.io/greasyfork/dt/562098?label=Downloads&color=blue)](https://greasyfork.org/zh-CN/scripts/562098)
[![Greasy Fork Version](https://img.shields.io/greasyfork/v/562098?label=Version&color=green)](https://greasyfork.org/zh-CN/scripts/562098)
![License](https://img.shields.io/github/license/peter-erer/buaa-spoc-helper?label=License)
专为北航学子打造的课程学习增强工具。一键导出“智学北航”课程回放的字幕与 PPT 讲义，支持智能笔记整理与完美打印排版。

## ✨ 核心功能

* **🎬 字幕导出 (SRT)**：生成标准 SRT 字幕文件，支持 PotPlayer 等播放器挂载，时间轴精准。
* **📝 智能笔记 (TXT)**：导出纯文本笔记。内置**智能断句算法**，自动合并破碎的短句，方便导入 Notion/Obsidian 或发送给 ChatGPT/Gemini 等大模型总结。
* **📊 PPT 讲义导出 (PDF)**：
    * 一键抓取课程所有 PPT 图片。
    * **智能预加载**：解决直接打印出现的“空白页”问题。
    * **完美打印适配**：适配 A4 横向排版，去除多余边框，100% 宽度铺满，保留页码与时间戳。

## 🚀 安装方法

1.  安装浏览器扩展 [Tampermonkey (油猴)](https://www.tampermonkey.net/)。
2.  前往 [Greasy Fork] (https://greasyfork.org/zh-CN/scripts/562098) 点击安装脚本。
3.  或者直接在 GitHub 点击 `raw` 文件安装。

## 📖 使用指南

### 1. 导出字幕/笔记
1.  打开任意一节“智学北航”的课程回放页面。
2.  点击页面右侧悬浮球 **“::: 学习助手 :::”**。
3.  点击 **“🎬 导出 SRT”** 或 **“📝 导出 TXT”** 即可下载。

### 2. 导出 PPT (PDF 讲义)
1.  确保你点击过左侧的 **“PPT”** 标签页（为了触发数据加载）。
2.  点击悬浮球上的 **“📊 导出 PPT (讲义)”**。
3.  脚本会自动弹窗并开始预加载图片（带有进度提示）。
4.  加载完成后会自动唤起打印窗口。
5.  **⚠️ 打印设置关键点**：
    * **目标打印机**：选择 “另存为 PDF”。
    * **布局**：推荐选择 **“横向”**。

## 🛠️ 本地开发

如果你想参与贡献或修改代码：

1.  Clone 本仓库到本地：
    ```bash
    git clone [https://github.com/你的用户名/你的仓库名.git](https://github.com/你的用户名/你的仓库名.git)
    ```
2.  在 Chrome 中开启 Tampermonkey 的 "允许访问文件网址" 权限。
3.  新建一个油猴脚本，通过 `@require` 引入本地文件进行调试：
    ```javascript
    // @require file://D:/path/to/your/script.js
    ```

## 🤝 贡献与反馈

欢迎提交 Issue 反馈 Bug 或建议新功能！欢迎 Pull Request！

## 📄 License

本项目基于 [MIT License](./LICENSE) 开源。仅供学习交流使用，请勿用于商业用途。