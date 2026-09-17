; Trình cài đặt Windows ban đầu cho SSBOM Manager
; Biên dịch bằng Inno Setup 6 sau khi chạy: py scripts/package_app.py
; Cài theo chuẩn D:\Sandbox\MP2027\huongdansetup_autoupdate.md

#define AppName "SSBOM Manager"
#define AppVersion "1.0.0"
#define AppPublisher "Kyocera Document Technology Vietnam - Production Engineering"
#define LauncherExe "SSBOM_Launcher.exe"

[Setup]
AppId={{D37E84B2-791A-4D78-9B21-889812A27F10}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Cài theo người dùng vì apps/<version>, current.json và .staging là trạng thái cập nhật
; có thể thay đổi, thuộc quyền sở hữu của người dùng thông thường không cần quyền admin.
DefaultDirName={localappdata}\SSBOM Manager
DefaultGroupName={#AppName}
OutputDir=..\release_artifacts
OutputBaseFilename=SSBOM_Manager_Setup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
UninstallDisplayName={#AppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\SSBOM_Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\current.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\update_sources.default.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\locales\*"; DestDir: "{app}\locales"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\apps\{#AppVersion}\*"; DestDir: "{app}\apps\{#AppVersion}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#LauncherExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#LauncherExe}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Tạo lối tắt trên màn hình nền"; GroupDescription: "Lối tắt bổ sung:"

[Run]
Filename: "{app}\{#LauncherExe}"; Description: "Khởi chạy {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\.staging"

