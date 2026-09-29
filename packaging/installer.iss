; Instalador de Windows (Inno Setup 6).
;
; Lo ejecuta scripts/build.py después de PyInstaller:
;     ISCC /DAppVersion=X.Y.Z /Odist packaging\installer.iss
;
; Se instala por usuario (no pide permisos de administrador). Al
; desinstalar borra siempre la caché de la ventana y pregunta si borrar
; también los datos de Documentos\Iberostar Gestor de Pedidos (por
; defecto, No).

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

[UninstallDelete]
; Caché de WebView2 (ver _webview_storage_dir en app/backend/desktop.py).
Type: filesandordirs; Name: "{localappdata}\{#AppName}"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep <> usPostUninstall then
    Exit;

  DataDir := ExpandConstant('{userdocs}\{#AppName}');

  if UninstallSilent or not DirExists(DataDir) then
    Exit;

  if MsgBox('¿Quieres borrar también tus datos?' + #13#10 + #13#10 +
            DataDir + #13#10 + #13#10 +
            'Se eliminarán los Excel mensuales, las plantillas y las copias ' +
            'de seguridad. Esta acción no se puede deshacer.',
            mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
    DelTree(DataDir, True, True, True);
end;
