#define MyAppName "MANAGEMENT CONSOLE"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Hamza Senhaji"
#define MyAppExeName "ManagementConsole.exe"

[Setup]
AppId={{MANAGEMENT-CONSOLE-HAMZA-2026}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\MANAGEMENT CONSOLE
DefaultGroupName=MANAGEMENT CONSOLE

OutputDir=installer_output
OutputBaseFilename=ManagementConsole-Setup

Compression=lzma
SolidCompression=yes

WizardStyle=modern
PrivilegesRequired=admin

UninstallDisplayIcon={app}\{#MyAppExeName}

[Files]
Source: "dist\ManagementConsole.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\MANAGEMENT CONSOLE"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\MANAGEMENT CONSOLE"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch MANAGEMENT CONSOLE"; Flags: nowait postinstall skipifsilent