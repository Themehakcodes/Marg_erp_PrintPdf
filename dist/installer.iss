#define MyAppName "Marg ERP Auto Printer"
#define MyAppVersion "1.0"
#define MyAppPublisher "TheMehakCodes"
#define MyAppDeveloper "Mehak Singh"
#define MyAppURL "https://themehakcodes.com"
#define MyAppExeName "marg_auto_printer.exe"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AppCopyright=Developed by {#MyAppDeveloper} @ {#MyAppPublisher}
DefaultDirName={autopf}\Marg ERP Auto Printer
DefaultGroupName={#MyAppName}
OutputDir=Output
OutputBaseFilename=Marg_ERP_Auto_Printer_Setup
SetupIconFile=logo.ico
WizardImageFile=logo.png
WizardSmallImageFile=logo.png
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoCopyright=Developed by {#MyAppDeveloper} @ {#MyAppPublisher}

; ============================================================
;  ADMIN PRIVILEGES — forces UAC prompt on launch
;  PrivilegesRequired=admin  → installer demands elevation
;  PrivilegesRequiredOverridesAllowed is intentionally omitted
;  so the user cannot bypass the UAC prompt
; ============================================================
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";  Description: "Create a &desktop shortcut";               GroupDescription: "Additional icons:"
Name: "startupicon";  Description: "Launch automatically at &Windows startup";  GroupDescription: "Startup:"

[Files]
Source: "marg_auto_printer.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "SumatraPDF.exe";        DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico";              DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}";           Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\logo.ico"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\logo.ico"; Tasks: desktopicon
Name: "{userstartup}\{#MyAppName}";     Filename: "{app}\{#MyAppExeName}"; Tasks: startupicon

[Run]
; Run the app after install — also elevated so it can write config.json to Program Files
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent runascurrentuser

; ======================================================
; CUSTOM WIZARD PAGES  —  PRINTER DROPDOWN + CONFIG
; ======================================================
[Code]

{ ------------------------------------------------------------------ }
{  Check for admin rights at startup — abort if not elevated          }
{  (belt-and-suspenders on top of PrivilegesRequired=admin)           }
{ ------------------------------------------------------------------ }
function InitializeSetup(): Boolean;
begin
  Result := True;
  if not IsAdminInstallMode then
  begin
    MsgBox(
      'This installer requires administrator privileges.' + #13#10 + #13#10 +
      'Please right-click the installer and select "Run as administrator".',
      mbCriticalError, MB_OK
    );
    Result := False;
  end;
end;

{ ------------------------------------------------------------------ }
{  Global variables                                                    }
{ ------------------------------------------------------------------ }
var
  PrinterPage  : TWizardPage;
  PrinterCombo : TComboBox;
  RefreshBtn   : TButton;

  FolderPage   : TInputDirWizardPage;
  PrefixPage   : TInputQueryWizardPage;
  IntervalPage : TInputQueryWizardPage;
  SilentPage   : TInputOptionWizardPage;

{ ------------------------------------------------------------------ }
{  Helper: Boolean -> JSON literal                                     }
{ ------------------------------------------------------------------ }
function BoolToJsonStr(B: Boolean): String;
begin
  if B then Result := 'true' else Result := 'false';
end;

{ ------------------------------------------------------------------ }
{  Helper: escape backslashes for JSON strings                         }
{ ------------------------------------------------------------------ }
function EscapeBackslashes(S: String): String;
begin
  StringChangeEx(S, '\', '\\', True);
  Result := S;
end;

{ ------------------------------------------------------------------ }
{  Helper: validate positive integer                                   }
{ ------------------------------------------------------------------ }
function IsPositiveInteger(S: String): Boolean;
var
  I, N: Integer;
begin
  Result := False;
  if Length(S) = 0 then Exit;
  N := 0;
  for I := 1 to Length(S) do
  begin
    if (S[I] < '0') or (S[I] > '9') then Exit;
    N := N * 10 + (Ord(S[I]) - Ord('0'));
  end;
  Result := (N > 0);
end;

{ ------------------------------------------------------------------ }
{  Populate PrinterCombo via PowerShell (primary) + WMIC (fallback)   }
{ ------------------------------------------------------------------ }
procedure PopulatePrinters;
var
  TempFile   : String;
  Lines      : TArrayOfString;
  I          : Integer;
  ResultCode : Integer;
  Line       : String;
begin
  PrinterCombo.Items.Clear;
  TempFile := ExpandConstant('{tmp}\printers_list.txt');

  { Primary: PowerShell Get-Printer }
  Exec(
    ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -NonInteractive -Command ' +
    '"Get-Printer | Select-Object -ExpandProperty Name | ' +
    'Out-File -Encoding ASCII -FilePath ''' + TempFile + '''"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode
  );

  if FileExists(TempFile) and LoadStringsFromFile(TempFile, Lines) then
    for I := 0 to GetArrayLength(Lines) - 1 do
    begin
      Line := Trim(Lines[I]);
      if Line <> '' then
        PrinterCombo.Items.Add(Line);
    end;

  { Fallback: WMIC }
  if PrinterCombo.Items.Count = 0 then
  begin
    Exec(
      ExpandConstant('{sys}\cmd.exe'),
      '/C wmic printer get Name /format:list > "' + TempFile + '"',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode
    );
    if FileExists(TempFile) and LoadStringsFromFile(TempFile, Lines) then
      for I := 0 to GetArrayLength(Lines) - 1 do
      begin
        Line := Trim(Lines[I]);
        if (Length(Line) > 5) and (Copy(Line, 1, 5) = 'Name=') then
          PrinterCombo.Items.Add(Trim(Copy(Line, 6, Length(Line) - 5)));
      end;
  end;

  if PrinterCombo.Items.Count > 0 then
    PrinterCombo.ItemIndex := 0
  else
  begin
    PrinterCombo.Items.Add('(No printers found – install a printer then click Refresh)');
    PrinterCombo.ItemIndex := 0;
  end;
end;

{ ------------------------------------------------------------------ }
{  Refresh button handler                                              }
{ ------------------------------------------------------------------ }
procedure OnRefreshClick(Sender: TObject);
begin
  PopulatePrinters;
  if PrinterCombo.Items.Count > 0 then
    MsgBox('Printer list refreshed. Found ' + IntToStr(PrinterCombo.Items.Count) +
           ' printer(s).', mbInformation, MB_OK)
  else
    MsgBox('No printers detected. Please install a printer in Windows first.',
           mbError, MB_OK);
end;

{ ------------------------------------------------------------------ }
{  Build all wizard pages                                              }
{ ------------------------------------------------------------------ }
procedure InitializeWizard;
var
  Lbl : TLabel;
begin
  { PAGE 1 – Printer Selection }
  PrinterPage := CreateCustomPage(wpWelcome,
    'Printer Configuration',
    'Select the printer for automatic PDF printing.');

  Lbl := TLabel.Create(WizardForm);
  with Lbl do begin
    Parent   := PrinterPage.Surface;
    Caption  := 'All printers currently installed on this computer are listed below.' + #13#10 +
                'Select the one that Marg ERP Auto Printer should send jobs to.';
    Left     := 0; Top := 0;
    Width    := PrinterPage.SurfaceWidth;
    WordWrap := True; AutoSize := True;
  end;

  Lbl := TLabel.Create(WizardForm);
  with Lbl do begin
    Parent := PrinterPage.Surface;
    Caption := 'Select Printer:';
    Left := 0; Top := 52; AutoSize := True;
  end;

  PrinterCombo := TComboBox.Create(WizardForm);
  with PrinterCombo do begin
    Parent    := PrinterPage.Surface;
    Left      := 0; Top := 70;
    Width     := PrinterPage.SurfaceWidth - 110;
    Style     := csDropDownList;
    Font.Size := 9;
  end;

  RefreshBtn := TButton.Create(WizardForm);
  with RefreshBtn do begin
    Parent  := PrinterPage.Surface;
    Caption := '↻  Refresh';
    Left    := PrinterCombo.Left + PrinterCombo.Width + 10;
    Top     := PrinterCombo.Top - 1;
    Width   := 95;
    Height  := PrinterCombo.Height + 2;
    OnClick := @OnRefreshClick;
  end;

  Lbl := TLabel.Create(WizardForm);
  with Lbl do begin
    Parent     := PrinterPage.Surface;
    Caption    := 'Tip: If your printer is missing, install it via Windows Settings → Printers & Scanners, then click Refresh.';
    Left       := 0;
    Top        := PrinterCombo.Top + PrinterCombo.Height + 12;
    Width      := PrinterPage.SurfaceWidth;
    WordWrap   := True; AutoSize := True;
    Font.Color := $00666666; Font.Size := 8;
  end;

  PopulatePrinters;

  { PAGE 2 – Watch Folder }
  FolderPage := CreateInputDirPage(PrinterPage.ID,
    'Watch Folder', 'Select Folder to Monitor',
    'Choose the folder where PDF files will be detected and automatically sent to the printer.',
    False, '');
  FolderPage.Add('');

  { PAGE 3 – File Prefix }
  PrefixPage := CreateInputQueryPage(FolderPage.ID,
    'File Prefix Filter', 'Set PDF File Prefix',
    'Only PDF files whose names begin with this prefix will be printed (e.g. MC_PRINT).' + #13#10 +
    'Leave blank to print ALL PDF files in the watch folder.');
  PrefixPage.Add('File Prefix:', False);
  PrefixPage.Values[0] := 'MC_PRINT';

  { PAGE 4 – Check Interval }
  IntervalPage := CreateInputQueryPage(PrefixPage.ID,
    'Check Interval', 'Set Folder Polling Interval',
    'How frequently (in seconds) should the app scan the watch folder for new PDF files?');
  IntervalPage.Add('Interval (seconds):', False);
  IntervalPage.Values[0] := '5';

  { PAGE 5 – Silent Mode }
  SilentPage := CreateInputOptionPage(IntervalPage.ID,
    'Silent Mode', 'Enable Silent / Background Printing',
    'When enabled, the application prints in the background without any dialogs or notifications.',
    False, False);
  SilentPage.Add('Enable Silent Mode (recommended for automated use)');
  SilentPage.Values[0] := True;
end;

{ ------------------------------------------------------------------ }
{  Validation on Next click                                            }
{ ------------------------------------------------------------------ }
function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;

  if CurPageID = PrinterPage.ID then
  begin
    if (PrinterCombo.Items.Count = 0) or
       (Pos('No printers found', PrinterCombo.Items[PrinterCombo.ItemIndex]) > 0) then
    begin
      MsgBox('Please select a valid printer before continuing.' + #13#10 +
             'If none appear, install a printer in Windows and click Refresh.',
             mbError, MB_OK);
      Result := False; Exit;
    end;
  end;

  if CurPageID = FolderPage.ID then
  begin
    if Trim(FolderPage.Values[0]) = '' then
    begin
      MsgBox('Please select a watch folder before continuing.', mbError, MB_OK);
      Result := False; Exit;
    end;
  end;

  if CurPageID = IntervalPage.ID then
  begin
    if not IsPositiveInteger(Trim(IntervalPage.Values[0])) then
    begin
      MsgBox('Please enter a valid positive whole number for the interval (e.g. 5).',
             mbError, MB_OK);
      Result := False; Exit;
    end;
  end;
end;

{ ------------------------------------------------------------------ }
{  Write config.json after install                                     }
{ ------------------------------------------------------------------ }
procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigFile      : String;
  ConfigText      : String;
  SelectedPrinter : String;
begin
  if CurStep = ssPostInstall then
  begin
    SelectedPrinter := '';
    if PrinterCombo.Items.Count > 0 then
      SelectedPrinter := PrinterCombo.Items[PrinterCombo.ItemIndex];

    ConfigFile := ExpandConstant('{app}\config.json');
    ConfigText :=
      '{'                                                                          + #13#10 +
      '    "printer": "'        + EscapeBackslashes(SelectedPrinter)     + '",'   + #13#10 +
      '    "watch_folder": "'  + EscapeBackslashes(FolderPage.Values[0])+ '",'   + #13#10 +
      '    "file_prefix": "'   + PrefixPage.Values[0]                    + '",'   + #13#10 +
      '    "check_interval": ' + Trim(IntervalPage.Values[0])             + ','   + #13#10 +
      '    "silent_mode": '    + BoolToJsonStr(SilentPage.Values[0])      + ','   + #13#10 +
      '    "developer": "Mehak Singh",'                                            + #13#10 +
      '    "company": "TheMehakCodes",'                                            + #13#10 +
      '    "version": "1.0"'                                                       + #13#10 +
      '}';

    SaveStringToFile(ConfigFile, ConfigText, False);
  end;
end;

{ ------------------------------------------------------------------ }
{  Finish page summary                                                 }
{ ------------------------------------------------------------------ }
procedure CurPageChanged(CurPageID: Integer);
var
  SelectedPrinter : String;
begin
  if CurPageID = wpFinished then
  begin
    SelectedPrinter := '(none selected)';
    if PrinterCombo.Items.Count > 0 then
      SelectedPrinter := PrinterCombo.Items[PrinterCombo.ItemIndex];

    WizardForm.FinishedLabel.Caption :=
      'Marg ERP Auto Printer has been successfully installed!'   + #13#10 + #13#10 +
      'Configuration saved:'                                     + #13#10 +
      '  Printer  :  ' + SelectedPrinter                        + #13#10 +
      '  Folder   :  ' + FolderPage.Values[0]                   + #13#10 +
      '  Prefix   :  ' + PrefixPage.Values[0]                   + #13#10 +
      '  Interval :  ' + IntervalPage.Values[0] + ' seconds'    + #13#10 + #13#10 +
      'Developed by Mehak Singh  |  TheMehakCodes'              + #13#10 +
      'https://themehakcodes.com';
  end;
end;