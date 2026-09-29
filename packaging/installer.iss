; Instalador de Windows (Inno Setup 6).
;
; Lo ejecuta scripts/build.py después de PyInstaller:
;     ISCC /DAppVersion=X.Y.Z /Odist packaging\installer.iss
;
; Se instala por usuario (no pide permisos de administrador) y al
; desinstalar NO borra los datos de Documentos\Iberostar Gestor de Pedidos.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#define AppName "Iberostar Gestor de Pedidos"
#define AppExe "IberostarGestorPedidos.exe"
#define AppPublisher "walteralee"
#define AppUrl "https://github.com/walteralee/iberostar-inventory-sync"

[Setup]
AppId={{6B7F3C2A-9E41-4D8B-A5C3-2F1E8D9B7A64}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
; Nombre fijo para que el enlace "releases/latest/download/..." no cambie.
OutputBaseFilename=IberostarGestorPedidos-Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\IberostarGestorPedidos\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
