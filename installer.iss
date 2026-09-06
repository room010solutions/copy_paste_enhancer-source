#define MyAppName "Clipboard Manager Pro"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Room 010 Solutions"
#define MyAppExeName "ClipboardManagerPro.exe"

[Setup]
; Application GUID
AppId={{7680b0dc-b7e2-4436-a14b-5e9b6a729d65}

AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}

DisableProgramGroupPage=yes
PrivilegesRequired=lowest

OutputDir=installer_output
OutputBaseFilename=ClipboardManagerPro-Setup-{#MyAppVersion}

Compression=lzma2
SolidCompression=yes

WizardStyle=modern

SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

CloseApplications=force
RestartApplications=no

; License Agreement
; The user must accept the license before continuing installation.
LicenseFile=license.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon"; Description: "Launch automatically when Windows starts"; GroupDescription: "Startup:"; Flags: unchecked

[Files]
; Main application
Source: "dist\ClipboardManagerPro.exe"; DestDir: "{app}"; Flags: ignoreversion

; License agreement - keep a copy inside the installed application folder
Source: "license.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent