;This file will be executed next to the application bundle image
;I.e. current directory will contain folder AutoMappingExcelProject with application files
[Setup]
AppId={{fxApplication}}
AppName=AutoMappingExcelProject
AppVersion=1.0.0
AppVerName=AutoMappingExcelProject 1.0.0
AppPublisher=bryan.yen
AppComments=AutoMappingExcelProject
AppCopyright=Copyright (C) 2018
;AppPublisherURL=http://java.com/
;AppSupportURL=http://java.com/
;AppUpdatesURL=http://java.com/
DefaultDirName={pf}\AutoMappingExcelProject
DisableStartupPrompt=Yes
DisableDirPage=No
DisableProgramGroupPage=Yes
DisableReadyPage=No
DisableFinishedPage=No
DisableWelcomePage=No
DefaultGroupName=bryan.yen
;Optional License
LicenseFile=
;WinXP or above
MinVersion=0,5.1 
OutputBaseFilename=AutoMappingExcelProject-1.0.0
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=AutoMappingExcelProject\AutoMappingExcelProject.ico
UninstallDisplayIcon={app}\AutoMappingExcelProject.ico
UninstallDisplayName=AutoMappingExcelProject
WizardImageStretch=No
WizardSmallImageFile=AutoMappingExcelProject-setup-icon.bmp   
ArchitecturesInstallIn64BitMode=x64


[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "AutoMappingExcelProject\AutoMappingExcelProject.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "AutoMappingExcelProject\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\AutoMappingExcelProject"; Filename: "{app}\AutoMappingExcelProject.exe"; IconFilename: "{app}\AutoMappingExcelProject.ico"; Check: returnTrue()
Name: "{commondesktop}\AutoMappingExcelProject"; Filename: "{app}\AutoMappingExcelProject.exe";  IconFilename: "{app}\AutoMappingExcelProject.ico"; Check: returnFalse()


[Run]
Filename: "{app}\AutoMappingExcelProject.exe"; Parameters: "-Xappcds:generatecache"; Check: returnFalse()
Filename: "{app}\AutoMappingExcelProject.exe"; Description: "{cm:LaunchProgram,AutoMappingExcelProject}"; Flags: nowait postinstall skipifsilent; Check: returnTrue()
Filename: "{app}\AutoMappingExcelProject.exe"; Parameters: "-install -svcName ""AutoMappingExcelProject"" -svcDesc ""AutoMappingExcelProject"" -mainExe ""AutoMappingExcelProject.exe""  "; Check: returnFalse()

[UninstallRun]
Filename: "{app}\AutoMappingExcelProject.exe "; Parameters: "-uninstall -svcName AutoMappingExcelProject -stopOnUninstall"; Check: returnFalse()

[Code]
function returnTrue(): Boolean;
begin
  Result := True;
end;

function returnFalse(): Boolean;
begin
  Result := False;
end;

function InitializeSetup(): Boolean;
begin
// Possible future improvements:
//   if version less or same => just launch app
//   if upgrade => check if same app is running and wait for it to exit
//   Add pack200/unpack200 support? 
  Result := True;
end;  
