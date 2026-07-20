; LeftHandDevice - Inno Setup installer definition
; Usage: scripts\build-installer.ps1

#ifndef MyAppVersion
  #define MyAppVersion "2.0.14"
#endif

#define MyAppName "LeftHandDevice"
#define MyAppPublisher "kazu-1234"
#define MyAppURL "https://github.com/kazu-1234/LeftHandDevice"
#define MyAppExeName "LeftHandDevice.exe"
#define PublishDir "..\dist\folder"

[Setup]
AppId={{C4A91E2B-6D38-4F5A-9B07-1E8C2A5D7F90}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE_NOTICE.txt
OutputDir=..\dist\installer
OutputBaseFilename=LeftHandDevice-v{#MyAppVersion}-win-x64-setup
SetupIconFile=..\WindowsApp\WinApp\Assets\AppIcon.ico
UninstallDisplayIcon={app}\Assets\AppIcon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
InfoAfterFile=
CloseApplications=force
RestartApplications=no
UsePreviousAppDir=yes

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE_NOTICE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\AppIcon.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\AppIcon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; アプリ本体と exe 隣の設定ファイルを削除（アップグレード時は走らない）
Type: filesandordirs; Name: "{app}"

[Code]
procedure TerminateApp;
var
  ResultCode: Integer;
  ExePath: String;
begin
  ExePath := ExpandConstant('{localappdata}\Programs\{#MyAppName}\{#MyAppExeName}');
  if FileExists(ExePath) then
    Exec(ExePath, '--exit', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#MyAppExeName}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(800);
end;

function InitializeSetup(): Boolean;
begin
  TerminateApp;
  Result := True;
end;

function InitializeUninstall(): Boolean;
begin
  TerminateApp;
  Result := True;
end;

// アップグレード時: exe 隣の設定を退避してから {app} を掃除し、インストール後に戻す
procedure BackupUserDataFile(const FileName: String);
var
  Src, DestDir, Dest: String;
begin
  Src := ExpandConstant('{app}\' + FileName);
  DestDir := ExpandConstant('{userappdata}\{#MyAppName}\upgrade-backup');
  Dest := DestDir + '\' + FileName;
  if FileExists(Src) then
  begin
    ForceDirectories(DestDir);
    CopyFile(Src, Dest, False);
  end;
end;

procedure RestoreUserDataFile(const FileName: String);
var
  Src, Dest: String;
begin
  Src := ExpandConstant('{userappdata}\{#MyAppName}\upgrade-backup\' + FileName);
  Dest := ExpandConstant('{app}\' + FileName);
  if FileExists(Src) then
  begin
    ForceDirectories(ExtractFilePath(Dest));
    CopyFile(Src, Dest, False);
    DeleteFile(Src);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    if DirExists(ExpandConstant('{app}')) then
    begin
      BackupUserDataFile('app_settings.json');
      BackupUserDataFile('app_patterns.json');
      BackupUserDataFile('saved_com_port.txt');
      DelTree(ExpandConstant('{app}'), True, True, True);
    end;
  end
  else if CurStep = ssPostInstall then
  begin
    RestoreUserDataFile('app_settings.json');
    RestoreUserDataFile('app_patterns.json');
    RestoreUserDataFile('saved_com_port.txt');
    DelTree(ExpandConstant('{userappdata}\{#MyAppName}\upgrade-backup'), True, True, True);
  end;
end;
