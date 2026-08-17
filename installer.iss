; =====================================================================
; Sanjivani Attendance System - Professional Inno Setup Script
; Configures unified multi-resolution icon across Installer, Uninstaller,
; Desktop Shortcut, and Start Menu.
; =====================================================================

#define MyAppName "Sanjivani Attendance"
#define MyAppVersion "2.0"
#define MyAppPublisher "Sanjivani AI"
#define MyAppExeName "Sanjivani Attendance.exe"
#define MyAppIcon "sanjivani.ico"

[Setup]
; App Identity
AppId={{D3E8F4A1-7B2C-4E90-881A-95E6B4F1A023}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://sanjivani.edu.in
AppSupportURL=https://sanjivani.edu.in
AppUpdatesURL=https://sanjivani.edu.in
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; Output Configuration
OutputDir=dist
OutputBaseFilename=Sanjivani-Attendance-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern

; Icon Settings for Installer & Uninstaller
SetupIconFile={#MyAppIcon}
UninstallDisplayIcon={app}\{#MyAppIcon}

; Permissions & Architecture
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ChangesAssociations=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Primary Executable
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Multi-Resolution Icon Assets
Source: "sanjivani.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "sanjivani.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.png"; DestDir: "{app}"; Flags: ignoreversion
; Static Haar Cascade XML Resource
Source: "haarcascade_frontalface_default.xml"; DestDir: "{app}"; Flags: ignoreversion
; Documentation
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
; Start Menu Shortcut with Explicit Icon Location
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"; WorkingDir: "{app}"
; Desktop Shortcut with Explicit Icon Location
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"; WorkingDir: "{app}"; Tasks: desktopicon

[InstallDelete]
; Purge old/stale shortcuts from Desktop and Programs to prevent generic icon caching
Type: files; Name: "{autodesktop}\Sanjivani-Attendance-System.lnk"
Type: files; Name: "{autodesktop}\{#MyAppName}.lnk"
Type: files; Name: "{autodesktop}\Face Attendance.lnk"
Type: files; Name: "{userdesktop}\Sanjivani-Attendance-System.lnk"
Type: files; Name: "{userdesktop}\{#MyAppName}.lnk"
Type: files; Name: "{userdesktop}\Face Attendance.lnk"
Type: files; Name: "{commondesktop}\Sanjivani-Attendance-System.lnk"
Type: files; Name: "{commondesktop}\{#MyAppName}.lnk"
Type: files; Name: "{autoprograms}\Sanjivani-Attendance-System.lnk"
Type: files; Name: "{autoprograms}\{#MyAppName}.lnk"
Type: files; Name: "{userprograms}\Sanjivani-Attendance-System.lnk"
Type: files; Name: "{userprograms}\{#MyAppName}.lnk"

[UninstallDelete]
Type: files; Name: "{autodesktop}\{#MyAppName}.lnk"
Type: files; Name: "{userdesktop}\{#MyAppName}.lnk"
Type: files; Name: "{autoprograms}\{#MyAppName}.lnk"
Type: files; Name: "{userprograms}\{#MyAppName}.lnk"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
procedure SHChangeNotify(wEventId: LongInt; uFlags: Cardinal; dwItem1, dwItem2: Cardinal);
external 'SHChangeNotify@shell32.dll stdcall';

// Notify Windows Shell to refresh icon cache immediately
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // SHCNE_ASSOCCHANGED ($08000000)
    SHChangeNotify($08000000, 0, 0, 0);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    SHChangeNotify($08000000, 0, 0, 0);
  end;
end;
