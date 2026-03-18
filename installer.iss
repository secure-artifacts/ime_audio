[Setup]
AppName=Voice IME (语音输入助手)
AppVersion=1.0.15
DefaultDirName={autopf}\Voice IME
DefaultGroupName=Voice IME
OutputDir=build\Installer
OutputBaseFilename=VoiceIME_Setup
SetupIconFile=src\app.ico
UninstallDisplayIcon={app}\voice_ime.exe
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startup"; Description: "开机自动启动 (Run at Startup)"; GroupDescription: "系统选项:"

[Files]
Source: "build\Release\voice_ime.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "scripts\*"; DestDir: "{app}\scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "voice_ime_template.ini"; DestDir: "{app}"; DestName: "voice_ime.ini"; Flags: ignoreversion

[Icons]
Name: "{group}\Voice IME"; Filename: "{app}\voice_ime.exe"
Name: "{autodesktop}\Voice IME"; Filename: "{app}\voice_ime.exe"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "VoiceIME"; ValueData: """{app}\voice_ime.exe"""; Tasks: startup

[Run]
Filename: "{app}\voice_ime.exe"; Description: "{cm:LaunchProgram,Voice IME}"; Flags: nowait postinstall skipifsilent