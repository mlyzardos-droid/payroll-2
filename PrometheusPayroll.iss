#define MyAppName "Prometheus Payroll"
#define MyAppVersion "1.1.0"
#define MyAppPublisher "Prometheus Payroll"
#define MyAppExeName "PrometheusPayroll.exe"

[Setup]
AppId={{46D5F875-6DC7-4F1A-9BC2-3DBBC6E21C5D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Prometheus Payroll
DefaultGroupName={#MyAppName}
OutputDir=dist
OutputBaseFilename=PrometheusPayroll-Setup-{#MyAppVersion}
SetupIconFile=PrometheusPayroll.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\PrometheusPayroll.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: postinstall nowait skipifsilent
