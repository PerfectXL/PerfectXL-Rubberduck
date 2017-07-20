#define appname "PerfectXL.VbaCodeAnalyzer.Host"
#define fileversion GetFileVersion(AddBackslash(SourcePath) + "..\" + appname + "\bin\Debug\" + appname + ".exe")

[Setup]
AllowUNCPath=false
AppendDefaultDirName=False
AppId={{f9fdde71-5b41-4a78-903c-5520ac7fdfe4}
AppName={#appname}
AppPublisher=Infotron B.V.
AppVersion={#fileversion}
Compression=lzma2/ultra64
DefaultDirName=C:\Program Files\{#appname}
DirExistsWarning=no
DisableDirPage=auto
DisableReadyPage=yes
LanguageDetectionMethod=none
OutputBaseFilename=setup-{#appname}-{#fileversion}
OutputDir=..\deploy
SetupLogging=true
ShowLanguageDialog=no
SolidCompression=true
SourceDir=..\{#appname}
VersionInfoCopyright=Copyright 2017 Infotron B.V.
VersionInfoVersion={#fileversion}

[Files]
Source: "bin\Debug\*.exe"; DestDir: "{app}"; Excludes: "*.vshost*"
Source: "bin\Debug\*.exe.config"; DestDir: "{app}"; Flags: onlyifdoesntexist confirmoverwrite uninsneveruninstall; Excludes: "*.vshost*"
Source: "bin\Debug\*.dll"; DestDir: "{app}"
Source: "bin\Debug\*.pdb"; DestDir: "{app}"
Source: "bin\Debug\license"; DestDir: "{app}"

[UninstallRun]
Filename: "{app}\{#appname}.exe"; Parameters: "stop"; WorkingDir: "{app}"; Flags: waituntilterminated
Filename: "{app}\{#appname}.exe"; Parameters: "uninstall"; WorkingDir: "{app}"; Flags: waituntilterminated

[UninstallDelete]
Type: files; Name: "{app}\*.exe"
Type: files; Name: "{app}\*.dll"
Type: files; Name: "{app}\*.pdb"
Type: files; Name: "{app}\license"
Type: dirifempty; Name: "{app}"

[Messages]
InstallingLabel=Installing [name] on this computer. Note that a username/password prompt appears at the top left of the screen.

[Code]
/////////////////////////////////////////////////////////////////////
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if (not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString)) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;


/////////////////////////////////////////////////////////////////////
function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;


/////////////////////////////////////////////////////////////////////
function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
// Return Values:
// 1 - uninstall string is empty
// 2 - error executing the UnInstallString
// 3 - successfully executed the UnInstallString

  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString();
  if (sUnInstallString <> '') then
  begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if (Exec(sUnInstallString, '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES', '', SW_HIDE, ewWaitUntilTerminated, iResultCode)) then
      Result := 3
    else
      Result := 2;
  end
  else
    Result := 1;
end;

/////////////////////////////////////////////////////////////////////
procedure CurStepChanged(CurStep: TSetupStep);
var
  sAppName: String;
  iResultCode: Integer;
begin
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
      UnInstallOldVersion();
  end;
  if (CurStep=ssPostInstall) then
  begin
    sAppName := WizardDirValue() + '\{#appname}.exe';
    if (Exec(sAppName, 'install --interactive', '', SW_HIDE, ewWaitUntilTerminated, iResultCode)) then
      Exec(sAppName, 'start', '', SW_HIDE, ewWaitUntilTerminated, iResultCode)
  end;
end;
