#define MyAppName "نظام تسليم الرخام الاحترافي"
#define MyAppNameEn "Marble Delivery Pro"
#define MyAppVersion "10.0.5"
#define MyAppExeName "MarbleDeliveryPro.exe"
#define BuildRoot "build\MarbleDeliveryPro"
#define OutputRoot "dist"

[Setup]
AppId={{A3D7D824-6AC6-4DCA-A0EE-B691A7A9F3E0}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=ProTech
VersionInfoVersion=10.0.5.0
VersionInfoCompany=ProTech
VersionInfoDescription=Professional marble delivery management and reporting system
VersionInfoProductName={#MyAppNameEn}
VersionInfoProductVersion={#MyAppVersion}
DefaultDirName={localappdata}\Programs\MarbleDeliveryPro
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputRoot}
OutputBaseFilename=Setup_MarbleDeliveryPro_v10.0.5
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableWelcomePage=no
CloseApplications=yes
CloseApplicationsFilter=MarbleDeliveryPro.exe
RestartApplications=no
SetupLogging=yes
UsePreviousAppDir=yes
UsePreviousTasks=yes
CreateUninstallRegKey=yes
ChangesEnvironment=no

[Languages]
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "إنشاء اختصار على سطح المكتب"; GroupDescription: "اختصارات إضافية:"

[Files]
Source: "{#BuildRoot}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "docs\دليل الاستخدام.md"; DestDir: "{app}\Documentation"; Flags: ignoreversion
Source: "docs\خطة التطوير الشاملة.md"; DestDir: "{app}\Documentation"; Flags: ignoreversion
Source: "docs\ملاحظات الإصدار.md"; DestDir: "{app}\Documentation"; Flags: ignoreversion
Source: "docs\دليل_التكامل_المؤسسي.md"; DestDir: "{app}\Documentation"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\دليل الاستخدام"; Filename: "{app}\Documentation\دليل الاستخدام.md"
Name: "{group}\إلغاء تثبيت {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "تشغيل {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  DataDir, BackupDir, DbFile, BackupFile, Stamp: string;
begin
  if CurStep <> ssInstall then
    Exit;

  if ExpandConstant('{param:SkipDataBackup|0}') = '1' then
    Exit;

  DataDir := ExpandConstant('{localappdata}\MarbleDeliveryPro');
  DbFile := AddBackslash(DataDir) + 'marble_delivery_v7.db';
  if not FileExists(DbFile) then
    Exit;

  BackupDir := AddBackslash(DataDir) + 'backups';
  ForceDirectories(BackupDir);
  Stamp := GetDateTimeString('yyyymmdd_hhnnss', '-', ':');
  BackupFile := AddBackslash(BackupDir) + 'pre_upgrade_{#MyAppVersion}_' + Stamp + '.db';
  if not CopyFile(DbFile, BackupFile, False) then
    Log('تعذر إنشاء نسخة قاعدة البيانات قبل الترقية: ' + BackupFile)
  else begin
    Log('تم إنشاء نسخة قاعدة البيانات قبل الترقية: ' + BackupFile);
    { WAL is copied with the DB, because SQLite may not have checkpointed yet. }
    if FileExists(DbFile + '-wal') then
      if not CopyFile(DbFile + '-wal', BackupFile + '-wal', False) then
        Log('تعذر نسخ سجل WAL إلى النسخة الاحتياطية');
    if FileExists(DbFile + '-shm') then
      if not CopyFile(DbFile + '-shm', BackupFile + '-shm', False) then
        Log('تعذر نسخ ملف SHM إلى النسخة الاحتياطية');
  end;
end;
