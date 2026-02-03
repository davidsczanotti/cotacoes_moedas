#ifndef MyAppName
#define MyAppName "cotacoes_moedas"
#endif
#ifndef MyAppDisplayVersion
#define MyAppDisplayVersion "0.1.0"
#endif
#ifndef MyAppVersion
#define MyAppVersion "0.1.0.0"
#endif
#ifndef MyAppPublisher
#define MyAppPublisher "davidsc"
#endif
#ifndef MyAppSupportUrl
#define MyAppSupportUrl "mailto:davidsc@zanotti.com.br"
#endif
#ifndef MyAppExeName
#define MyAppExeName "cotacoes-moedas.exe"
#endif
#ifndef MyAppSourceDir
#define MyAppSourceDir "..\\dist\\main.dist"
#endif
#ifndef MyAppIcon
#define MyAppIcon "..\\imagem_ico\\finance.ico"
#endif

[Setup]
AppId={{9EAE9A2C-69B8-4E9E-8C9F-1D9F6C2C8C44}
AppName={#MyAppName}
AppVersion={#MyAppDisplayVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppSupportUrl}
AppSupportURL={#MyAppSupportUrl}
AppUpdatesURL={#MyAppSupportUrl}
AppContact={#MyAppSupportUrl}
DefaultDirName={pf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\\dist
OutputBaseFilename=cotacoes-moedas-setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#MyAppIcon}
UninstallDisplayIcon={app}\\{#MyAppExeName}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
UsePreviousAppDir=yes

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\\*"; DestDir: "{app}"; Flags: recursesubdirs

[Icons]
Name: "{group}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"
Name: "{userdesktop}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "Executar {#MyAppName}"; Flags: nowait postinstall skipifsilent
