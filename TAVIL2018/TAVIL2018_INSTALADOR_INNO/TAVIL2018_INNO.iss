; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "TAVIL2018"
#define MyAppNameDir "TAVIL2018.bundle"
#define MyAppVersion "2018.0.0.13"
#define MyAppPublisher "Copyright © Jose Alberto Torres (2aCAD Global Group  2018)"
#define MyAppURL "http://www.2acad.es"
#define MyWeb "2aCAD Global Group"
#define MyOrigen "{srcexe}\..\..\"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{30B097C4-D02B-4887-9ECF-93B603012737}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName="{commonappdata}\Autodesk\ApplicationPlugins\{#MyAppNameDir}"
DisableDirPage=yes
DefaultGroupName=2aCAD Global Group\TAVIL2018
DisableProgramGroupPage=yes
OutputDir="{srcexe}\..\SALIDA"
OutputBaseFilename=TAVIL2018_v{#MyAppVersion}
SetupIconFile=..\Resources\TAVIL_ico.ico
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={uninstallexe}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany=2aCAD Global Group
VersionInfoDescription=TAVIL2018. Administración de planos de Layout
VersionInfoCopyright=Copyright © Jose Alberto Torres (2aCAD Global Group  2018)
VersionInfoProductName=TAVIL2018
VersionInfoProductVersion={#MyAppVersion}
PrivilegesRequired=none

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Files]
; *** APP principal y fichero configuracion
Source: "{#MyOrigen}PackageContents.xml"; DestDir: "{app}"; DestName: "PackageContents.xml"; Flags: ignoreversion 
Source: "{#MyOrigen}..\..\{#MyAppName}_DOCUMENTOS\LAYOUTDBS4.xlsx"; DestDir: "{app}"; Flags: ignoreversion  
Source: "{#MyOrigen}{#MyAppName}.ini"; DestDir: "{app}"; Flags: ignoreversion  
Source: "{#MyOrigen}{#MyAppName}.ico"; DestDir: "{app}"; Flags: ignoreversion  
Source: "{#MyOrigen}bin\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyOrigen}bin\*.config"; DestDir: "{app}"; Flags: ignoreversion 
Source: "{#MyOrigen}bin\*.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyOrigen}Resources\{#MyAppName}.cuix"; DestDir: "{app}"; Flags: ignoreversion
; *** RESOURCES (Bloques, Excel, etc.)
Source: "{#MyOrigen}Resources\BloqueRecursos.dwg"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyOrigen}Resources\TAVIL.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyOrigen}bin\BLOQUES\*"; Excludes: "*.bak*,*.ac$"; DestDir: "{app}\BLOQUES"; Flags: ignoreversion createallsubdirs recursesubdirs; Permissions: everyone-full

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#MyWeb}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{group}\{#MyAppName}.dll.config"; Filename: "{app}\{#MyAppName}dll.config"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppName}.dll.ini"
Name: "{group}\TAVIL2018"; Filename: "{app}"; WorkingDir: "{app}"; IconFilename: "{app}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\BLOQUES\*.*"
Type: filesandordirs; Name: "{app}\*.*"
Type: filesandordirs; Name: "{app}\"
Type: filesandordirs; Name: "{commonappdata}\Autodesk\ApplicationPlugins\{#MyAppNameDir}"

[Dirs]
Name: "{app}\BLOQUES"; Flags: uninsalwaysuninstall

[Registry]
Root: "HKCU"; Subkey: "Software\Autodesk\AutoCAD\R22.0\ACAD-0001:40A\Profiles\<<Perfil sin nombre>>\Variables"; ValueType: string; ValueName: "TRUSTEDPATHS"; ValueData: "C:\\ProgramData\\Autodesk\\ApplicationPlugins\\TAVIL2018.bundle"; Flags: createvalueifdoesntexist uninsclearvalue; Permissions: everyone-full
Root: "HKLM"; Subkey: "SOFTWARE\Autodesk\AutoCAD\R22.0\ACAD-1001\Variables\TRUSTEDPATHS"; ValueType: string; ValueName: "@"; ValueData: "C:\\ProgramData\\Autodesk\\ApplicationPlugins\\TAVIL2018.bundle"; Flags: uninsclearvalue createvalueifdoesntexist

[Run]
;Filename: "{sys}\icacls.exe"; Parameters: "{app} /grant Everyone:F /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
;Filename: "{sys}\icacls.exe"; Parameters: "{app}\*.* /grant Everyone:F /t"; Flags: runmaximized; Description: "Dar permisos a todos los ficheros"
;** PERMISOS PARA TODOS Y USUARIO ACTUAL A LA CARPETA (Carpeta, Subcarpeta y ficheros - (OI)(CI)
Filename: "{sys}\icacls.exe"; Parameters: "{app} /grant:r Todos:(OI)(CI)F /inheritance:d /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
Filename: "{sys}\icacls.exe"; Parameters: "{app} /grant:r Usuarios:(OI)(CI)F /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
Filename: "{sys}\icacls.exe"; Parameters: "{app} /grant:r {username}:(OI)(CI)F /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
;** PERMISOS PARA TODOS Y USUARIO ACTUAL A LOS FICHEROS (Por si no está habilitada la herencia) (Carpeta, Subcarpeta y ficheros - (OI)(CI)
Filename: "{sys}\icacls.exe"; Parameters: "{app}\*.* /inheritance:d /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
Filename: "{sys}\icacls.exe"; Parameters: "{app}\* /T /reset"; Flags: runmaximized; Description: "Habilitar herencia en todas sub y files"
;Filename: "{sys}\icacls.exe"; Parameters: "{app}\*.* /grant Usuarios:(OI)(CI)F /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"
;Filename: "{sys}\icacls.exe"; Parameters: "{app}\*.* /grant {username}:(OI)(CI)F /t"; Flags: runmaximized; Description: "Dar permisos a la carpeta entera"

[ThirdParty]
UseRelativePaths=True
