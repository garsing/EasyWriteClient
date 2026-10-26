; 易写 Desktop 安装包。由 build-setup.ps1 调用 ISCC 编译。
; 装到当前用户目录，因为程序会把 config.json 写在 exe 旁边。

#define MyAppName "易写"
#define MyAppVersion "1.0.7"
#define MyAppExeName "EasyWriteClient.Desktop.exe"
#define MyAppPublisher "易写"

[Setup]
AppId={{8F4C2A1E-6B7D-4E91-A3C5-9D0E1F2A4B68}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion=1.0.7.0
VersionInfoProductName={#MyAppName}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=易写-Setup
SetupIconFile=..\Assets\yi-write.ico
UninstallDisplayIcon={app}\Assets\yi-write.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=EasyWriteClient.Desktop.exe,YiWriteBrowserBridge.exe
ArchitecturesAllowed=x86 x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加选项:"; Flags: checkedonce

[Files]
Source: "..\bin\Release\net472\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.pdb"
Source: "redist\MicrosoftEdgeWebview2Setup.exe"; Flags: dontcopy

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\yi-write.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Assets\yi-write.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动{#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function IsDotNet472Installed: Boolean;
var
  ReleaseVer: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', ReleaseVer) and (ReleaseVer >= 461808);
end;

function IsWebView2Installed: Boolean;
var
  Version: String;
begin
  Result :=
    (RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version) or
     RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version) or
     RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version));
  if Result then
    Result := (Version <> '') and (Version <> '0.0.0.0');
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  Result := '';
  if not IsDotNet472Installed then
  begin
    Result := '.NET Framework 4.7.2 未安装。请先安装后再运行本安装包。';
    exit;
  end;

  if IsWebView2Installed then
    exit;

  ExtractTemporaryFile('MicrosoftEdgeWebview2Setup.exe');

  if not Exec(ExpandConstant('{tmp}\MicrosoftEdgeWebview2Setup.exe'), '/silent /install', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
  begin
    Result := '无法启动 WebView2 安装程序。';
    exit;
  end;

  if ResultCode = 3010 then
    NeedsRestart := True
  else if ResultCode <> 0 then
    Result := '安装 WebView2 运行时失败，代码 ' + IntToStr(ResultCode) + '。';
end;
