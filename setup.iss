[Setup]
; Información básica de tu aplicación
AppName=Argentina RF Tool
AppVersion=1.0.3
AppPublisher=Manuel Alejandro Nuñez Ayala
DefaultGroupName=Argentina RF Tool

; CRÍTICO: La instalación debe ser en la carpeta del usuario (AppData) para no pedir permisos de Admin al auto-actualizar.
DefaultDirName={localappdata}\Programs\ArgentinaRFTool
PrivilegesRequired=lowest

; Dónde se guardará el instalador final y cómo se llamará
OutputDir=D:\ARGENTINA\MACRO TOOLS\ARGENTINA RF TOOL\NEW VERSION\APP\Instalador
OutputBaseFilename=SETUP_ArgentinaRFTool
SetupIconFile=D:\ARGENTINA\MACRO TOOLS\ARGENTINA RF TOOL\NEW VERSION\APP\argentina.ico
Compression=lzma
SolidCompression=yes

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Asegúrate de apuntar a la ruta exacta donde PyInstaller generó tu .exe
Source: "D:\ARGENTINA\MACRO TOOLS\ARGENTINA RF TOOL\NEW VERSION\APP\dist\Argentina RF Tool.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Argentina RF Tool"; Filename: "{app}\Argentina RF Tool.exe"
Name: "{autodesktop}\Argentina RF Tool"; Filename: "{app}\Argentina RF Tool.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ArgentinaRFTool.exe"; Description: "{cm:LaunchProgram,Argentina RF Tool}"; Flags: nowait postinstall skipifsilent