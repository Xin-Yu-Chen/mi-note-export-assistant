# 第三方开源组件声明

本项目（Mi Note Export Assistant）在自身代码之外，使用了以下开源软件。版权与许可均归各自所有者。

## 运行时依赖

| 组件 | 用途 | 许可证 |
|---|---|---|
| [Playwright for Python](https://github.com/microsoft/playwright-python) | 自动驱动系统浏览器、登录与读取 | Apache License 2.0 |
| [python-docx](https://github.com/python-openxml/python-docx) | 生成 Word（.docx） | MIT License |
| [Pillow](https://github.com/python-pillow/Pillow) | 图片处理 | HPND License（MIT 风格） |
| [Requests](https://github.com/psf/requests) | HTTP 请求（遗留渲染脚本） | Apache License 2.0 |
| Tkinter | 图形界面 | Python Software Foundation License（随 Python 内置） |
| SQLite | 本地数据库 | Public Domain（随 Python 内置） |

## 打包工具

| 组件 | 用途 | 许可证 |
|---|---|---|
| [PyInstaller](https://pyinstaller.org/) | 将程序打包成独立 exe | GPL（带 bootloader 特殊例外），允许打包/分发商业与非商业程序 |

PyInstaller 自身的许可证包含一个特殊例外条款，明确使用其打包生成的可执行文件不受 GPL 传染，可自由分发。

## 浏览器

打包后的程序默认调用 Windows 系统自带的 **Microsoft Edge**（Chromium 内核），**不随附、不分发浏览器二进制**；Edge 的使用受 Microsoft 相应条款约束。

## 关于商标

“小米”“红米”“Redmi”及小米相关产品名称、Logo 均为小米科技有限责任公司的商标或注册商标。本项目为独立的非官方工具，不主张任何相关权利，也不代表或暗示与小米存在关联或授权关系。

---

如你在本项目基础上进行分发或再发布，请保留本文件与各依赖的原始许可声明。完整许可证文本可从各项目官网或其代码仓库获取。
