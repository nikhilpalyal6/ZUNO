
#define MyAppName "ZUNO"
#define MyAppVersion "1.0.0"
#define MyAppExeName "ZUNO.exe"

[Setup]
AppId={{A6B3E29C-4E17-4B7F-9C62-82F5C13D7A41}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\ZUNO
DefaultGroupName=ZUNO
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=src\SmartShift.App\Assets\ZUNO.ico
OutputDir=artifacts\installer
OutputBaseFilename=ZUNO-Setup-1.0.0
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "Create a Desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "artifacts\ZUNO-win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ZUNO"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\ZUNO"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch ZUNO"; Flags: postinstall nowait skipifsilent
