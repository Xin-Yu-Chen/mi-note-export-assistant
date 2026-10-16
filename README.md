# 小米 / 红米笔记导出助手（Mi Note Export Assistant）

> ⚠️ **非官方社区工具**：本项目与小米公司 / 红米（Redmi）**没有任何关联、未获其授权或背书**。它只是自动化你本人在浏览器里就能做的操作，用于**备份你自己的**笔记。

一个运行在 Windows 上的桌面小工具，自动登录小米云、把**你本人账号**里的笔记**增量同步**到本地并导出 Word。无需打开 F12、无需粘贴代码：**扫码登录一次 → 点「立即同步」→ 点「导出 Word」**。

- 图形界面：登录 / 立即同步（增量）/ 导出 Word 三个按钮，带进度条与日志。
- **增量同步**：本地 SQLite 存全量，每次只拉“新增 / 修改”的笔记，删除的本地标记保留，每天同步通常只需几秒。
- **开箱即用**：双击运行；自动调用系统自带的 **Edge**，无需另装浏览器。
- 图片随笔记一起保存，导出 Word 时内嵌。
- 全程**本地运行、只连官方 `i.mi.com`、只读不写**，不改/删云端内容，数据不出本机。

---

## 它是怎么工作的（请先理解）

它**不是破解工具，也不接触小米服务器的漏洞**，而是用 [Playwright](https://playwright.dev/python/) 自动驱动你电脑上的 Edge，替你完成“登录 → 打开云笔记网页 → 逐条读取”这套**你本人手动也能做**的流程：

1. 启动一个独立的浏览器窗口（登录态保存在本地 `.browser-profile`，仅用于本工具）；
2. 你扫码登录自己的小米账号；
3. 先拉一份轻量清单（编号 + 修改时间），与本地数据库比对，只对新增/改动的笔记拉全文和图片；
4. 数据存到本地 `mi_notes.db`（SQLite），再由本地脚本排版生成 Word。

它**不保存、不上传你的账号密码**，也没有任何作者的服务器；你登录的是小米官方页面。

---

## 隐私与安全

- 数据库、图片、导出的 Word、登录态**全部保存在程序所在文件夹**，不经过任何第三方服务器。
- 程序只与 `i.mi.com` 官方域名通信；除登录和读取你自己的笔记外不做其它网络请求。
- 仓库的 `.gitignore` 已默认排除 `.browser-profile/`、`mi_notes*.db`、`*.json`、`*.docx`、`cookie.txt`，**请不要把这些文件提交到任何公开仓库或分享给他人**。

---

## 合规使用条款（重要）

下载、运行或二次分发即表示你同意：

1. **只用于备份、导出你本人拥有合法访问权的账号中的笔记**。不得用于抓取他人数据、批量采集、转售数据或任何违反小米账号 / 云服务协议的用途。
2. 不得将本工具用于商业数据采集、爬虫服务、账号代操作、对外售卖抓取能力等。
3. 访问小米云服务仍受[小米账号用户协议](https://account.xiaomi.com/)及相关条款约束；因你使用本工具产生的账号风控、封禁或法律责任由你自行承担。
4. 作者不对小米接口变更、数据丢失、账号异常、服务中断或任何直接/间接损失负责。**使用前请自行做好数据备份。**
5. 请勿在本工具名称、图标或发布内容中使用小米 / 红米的官方商标、Logo 或暗示官方授权的措辞。
6. 本工具按 “AS IS”（现状）提供，作者无持续维护、提供服务或保证可用的义务。

> 本项目开源、免费、仅用于个人学习与个人数据备份。如果你不确定某种用途是否合法合规，请不要使用，或咨询专业法律人士。

---

## 下载即用（Windows）

到右侧 **[Releases](../../releases)** 下载 `MiNoteExportAssistant-v1.0.0-windows-x64.exe`（单文件便携版，下载后可自行重命名为「小米笔记助手.exe」），双击运行即可，**无需安装 Python**。

> 本项目**没有购买代码签名证书**，因此首次运行可能出现 Windows 保护提示，这是未签名开源软件的常见现象：
>
> - **SmartScreen**：点“**更多信息**” → “**仍要运行**”。
> - **浏览器提示“不常下载”**：选择“保留” → “仍然保留”。
> - **杀毒软件误报**：PyInstaller 单文件程序可能被误报，可加信任或改用“从源码运行”（代码完全可审计）。
>
> 想验证文件未被篡改：每个 Release 都附 `SHA256SUMS.txt`，下载后在该目录执行
> `certutil -hashfile MiNoteExportAssistant-v1.0.0-windows-x64.exe SHA256`，与文件内的值比对（若已重命名，换成实际文件名即可）。

### 三个按钮

| 按钮 | 作用 |
|---|---|
| ① 登录小米账号 | 弹出 Edge，扫码登录一次，登录态保留几天 |
| ② 立即同步（增量） | 只拉新增/修改的笔记，首次会全量 |
| ③ 导出 Word | 用本地数据生成 Word（含封面、目录、分组、内嵌图片） |

建议把 exe 放在一个**固定、可写的文件夹**（如 `D:\我的笔记\`），数据与导出结果都在那里生成。

---

## 从源码运行（适合不信任 exe、想审计代码的人）

需要 Windows 10/11 与 [Python 3.10+](https://www.python.org/downloads/)。

```bat
git clone https://github.com/Xin-Yu-Chen/mi-note-export-assistant.git
cd mi-note-export-assistant
py -m pip install -r requirements.txt
py mi_note_app.py
```

或直接双击 `①安装依赖.bat` 再双击 `③打开助手（图形界面）.bat`。

---

## 自己打包成 exe

双击 **`⑤打包发布.bat`**（或运行 `py tools/make_release.py`）即可一键完成：
1. 用 PyInstaller 打包（产物 `dist\小米笔记助手.exe`）；
2. 若装了 [Inno Setup 6](https://jrsoftware.org/isdl.php)，再编译出带快捷方式/卸载器的安装包；
3. 把产物收集到 `release/`，并生成 **`SHA256SUMS.txt`**（发布到 Release 时一并上传，供用户校验，替代代码签名）。

只想要 exe 也可手动：

```bat
py -m pip install pyinstaller
py -m PyInstaller mi_note_app.spec --noconfirm
```

- 浏览器默认调用系统 Edge，无需打包 Chromium。
- 图标由 `tools/make_icon.py` 用 Pillow 现场绘制（打包时自动生成，不入库）；exe 版本信息在 `assets/version_info.txt`。
- **数据存哪里**：便携版（把 exe 放在普通文件夹）数据在 exe 旁；安装版（装到 Program Files 等只读位置）数据自动落到 `%APPDATA%\MiNoteExportAssistant`。

---

## 常见问题

- **又弹出登录页 / 401**：登录态过期（一般几天），点“① 登录”重新扫码即可。
- **提示找不到浏览器**：系统缺少 Edge 时可执行 `py -m playwright install chromium`，或安装 Edge。
- **换账号**：删除程序目录下的 `.browser-profile` 文件夹后重新登录。
- **私密笔记 / 待办 / 思维导图**：当前版本不导出，仅处理普通图文笔记。
- **Word 目录页码不对**：Word 中按 `Ctrl+A` 再按 `F9` 更新目录。

---

## 项目结构

```
mi_note_app.py        图形界面（Tkinter）
mi_note_core.py       核心：本地数据库 + Playwright 登录 + 增量同步 + 导出
export_mi_notes.py    Word 排版渲染（也支持 --from-json 从 JSON 生成）
mi_note_app.spec      PyInstaller 打包配置（含图标/版本信息）
tests/                核心逻辑测试（py tests/test_core.py）
tools/make_release.py 一键打包 + 安装包 + SHA256 校验
tools/make_icon.py    生成应用图标（打包时自动运行）
installer/            Inno Setup 安装包脚本
assets/               版本信息等资源（图标由脚本生成、不入库）
```

## 开源声明（第三方组件）

本项目使用了以下开源组件，版权归各自所有者，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)：
Playwright（Apache-2.0）、python-docx（MIT）、Pillow（HPND/MIT）、Requests（Apache-2.0）、Tkinter（PSF）、PyInstaller（含 bootloader 例外条款）。

## 许可证

[MIT License](LICENSE)（仅覆盖本项目代码；访问小米服务仍受其条款与法律约束）。
