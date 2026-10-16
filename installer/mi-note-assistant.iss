; ============================================================================
; 小米笔记导出助手 · Inno Setup 安装包脚本
; 前置：先用 PyInstaller 生成 dist\小米笔记助手.exe
; 编译：安装 Inno Setup 6（https://jrsoftware.org/isdl.php）后，
;       双击本文件并用其编译器打开，或运行 tools\make_release.py 自动编译。
;
; 说明：程序安装到 Program Files（只读），运行时数据会自动写入
;       %APPDATA%\MiNoteExportAssistant；卸载不会删除这些数据。
; ============================================================================

#define MyAppName "小米笔记导出助手"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "非官方开源项目"
#define MyAppExeName "小米笔记助手.exe"
#define MyAppId "{{BF4EEA6B-C39F-47C1-A588-F3D38FFFEF7C}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppComments=非官方工具，仅用于备份本人小米/红米笔记
DefaultDirName={autopf}\MiNoteExportAssistant
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=Setup_小米笔记助手_v1.0.0
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
; 不收集任何信息，无联网安装
CloseApplications=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "在桌面创建图标"; GroupDescription: "附加图标："

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
