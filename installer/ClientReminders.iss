; Installer for Client Reminders. Build the program first (PyInstaller, output in dist\ClientReminders),
; then compile this file with Inno Setup (ISCC.exe). The setup file is written to installer_output\.
; Installs for the current user only: no administrator rights needed.

[Setup]
AppId={{F930C728-E324-4CD5-B5ED-7CA60C616616}
AppName=Client Email Reminders
AppVersion=1.15
AppPublisher=PHP
DefaultDirName={localappdata}\Programs\ClientReminders
DefaultGroupName=Client Email Reminders
PrivilegesRequired=lowest
OutputDir=..\installer_output
OutputBaseFilename=ClientRemindersSetup
Compression=lzma2
SolidCompression=yes
UninstallDisplayName=Client Email Reminders
SetupIconFile=app_icon.ico

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
; The program and its packages. Data is NOT installed here: it lives in %LOCALAPPDATA%\ClientRemindersData
; and is kept when the program is updated or uninstalled.
Source: "..\dist\ClientEmailReminders\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Client Email Reminders"; Filename: "{app}\ClientEmailReminders.exe"
Name: "{group}\Uninstall Client Email Reminders"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Client Email Reminders"; Filename: "{app}\ClientEmailReminders.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ClientEmailReminders.exe"; Description: "Start Client Email Reminders"; Flags: nowait postinstall skipifsilent

[Code]
// Versions 1.1 and 1.2 were installed without an app ID, so Windows does not track them as installed apps.
// Before installing, this removes any such version, so no duplicate entries or old shortcuts are left behind.
// Your data (%LOCALAPPDATA%\ClientRemindersData) is never touched.
const
  CurrentAppKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{F930C728-E324-4CD5-B5ED-7CA60C616616}_is1';

function InitializeSetup(): Boolean;
var
  OldUninstaller: String;
  ResultCode: Integer;
begin
  Result := True;
  OldUninstaller := ExpandConstant('{localappdata}\Programs\ClientReminders\unins000.exe');

  // Only act when this app is not already registered (an update from 1.3 onward is handled by Windows itself)
  if (not RegKeyExists(HKEY_CURRENT_USER, CurrentAppKey)) and FileExists(OldUninstaller) then
  begin
    if MsgBox('An earlier version of Client Email Reminders is installed. It will be removed before this version is installed. Your data is kept. Continue?', mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
      Exit;
    end;

    if not Exec(OldUninstaller, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      MsgBox('The earlier version could not be removed. Remove it from Installed apps, then run this setup again.', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    // Remove the shortcuts the earlier version left behind
    DeleteFile(ExpandConstant('{userdesktop}\Client Reminders.lnk'));
    DelTree(ExpandConstant('{userprograms}\Client Reminders'), True, True, True);
  end;
end;

















