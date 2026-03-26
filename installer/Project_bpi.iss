#define AppName "Project_bpi"
#define AppVersion "1.0.0"
#define AppPublisher "Project_bpi"
#define AppExeName "Project_bpi.exe"
#define ProjectRoot ".."
#define ReleaseOutputDir ProjectRoot + "\Project_bpi\bin\Release"
#define DebugOutputDir ProjectRoot + "\Project_bpi\bin\Debug"

#ifndef BuildOutputDir
  #ifexist ReleaseOutputDir + "\" + AppExeName
    #define BuildOutputDir ReleaseOutputDir
  #else
    #define BuildOutputDir DebugOutputDir
  #endif
#endif

[Setup]
AppId={{83EBADCA-9DB6-4A1D-A1E7-30ABE93731A8}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=Project_bpi_Setup_{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#AppExeName}

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные значки:"

[Files]
Source: "{#BuildOutputDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.pdb,*.xml,*.cache,*.tmp,*.log,*.vshost.*,SavedTemplates\*,TemplateDatabases\*"

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Запустить {#AppName}"; Flags: nowait postinstall skipifsilent
