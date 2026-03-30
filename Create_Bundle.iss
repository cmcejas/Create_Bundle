; Inno Setup script for Create Bundle - PDF Bundler
; Build the exe first: pyinstaller Create_Bundle.spec
; Then compile this script in ISCC (Inno Setup Compiler).
; No admin rights required - installs per user.

#define MyAppName "Create Bundle"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Create Bundle"
#define MyAppExeName "Create_Bundle.exe"
#define MyAppOutputName "Create_Bundle_Setup"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Per-user install: no admin required. Installs under user's AppData\Local.
DefaultDirName={userappdata}\Create Bundle
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename={#MyAppOutputName}
SetupIconFile=
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; No admin - install for current user only
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=no
DisableWelcomePage=no
WizardImageFile=
WizardSmallImageFile=

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Visual C++ Runtime (required for PyInstaller bundles with Windows extensions)
Source: "Microsoft Visual C++ 14 Runtime"; Check: not VCRedistInstalled; DestDir: "{tmp}"; Flags: deleteafterinstall ignoreversion skipifsourcedoesntexist
; Single exe from onefile build (all dependencies inside the exe)
Source: "dist\Create_Bundle.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\INPUT"
Name: "{app}\OUTPUT"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "Merge PDFs, Word docs and Outlook emails into one PDF"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Comment: "Merge PDFs, Word docs and Outlook emails into one PDF"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}\INPUT"
Type: dirifempty; Name: "{app}\OUTPUT"

[Code]
function VCRedistInstalled: Boolean;
begin
  Result := FileExists('C:\Windows\System32\vcruntime140.dll') or 
            FileExists('C:\Windows\SysWOW64\vcruntime140.dll');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    Log('Create Bundle installed. INPUT and OUTPUT folders created in ' + ExpandConstant('{app}'));
end;
