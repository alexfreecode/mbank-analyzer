; installer.iss — skrypt Inno Setup budujący instalator mBank Analyzer
;
; Budowanie instalatora:
;   1) Najpierw uruchomić build_exe.ps1 — powstanie dist\mBank Analyzer.exe
;   2) Następnie skompilować ten plik kompilatorem Inno Setup (ISCC.exe installer.iss)
;
; Wynik: dist\mBank Analyzer Setup.exe — jeden plik do przekazywania użytkownikom.
;
; Instalacja przebiega w folderze profilu użytkownika (bez uprawnień
; administratora), tworzony jest skrót w menu Start i (opcjonalnie) na
; pulpicie, program rejestruje się w „Dodaj/usuń programy” z pełną
; obsługą odinstalowania.

#define MyAppName "mBank Analyzer"
#define MyAppVersion "1.1"
; Neutralna nazwa wydawcy — żadnych danych osobistych w instalatorze
; (program ma być swobodnie przekazywalny, bez niczyich danych osobowych)
#define MyAppPublisher "mBank Analyzer"
#define MyAppExeName "mBank Analyzer.exe"

[Setup]
AppId={{B6F2B6B0-2F0A-4C7E-9C9A-7E6B6E1A9D3A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
; Instalujemy bez uprawnień administratora — w folderze bieżącego użytkownika,
; aby niedoświadczony użytkownik nie natknął się na monit UAC
PrivilegesRequired=lowest
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=mBank Analyzer Setup
SetupIconFile=assets\app.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; Nazwa skrótu na pulpicie wyliczana jest dynamicznie (patrz [Code] niżej) —
; aby nie nadpisać już istniejącego skrótu o tej samej nazwie, który mógł
; zostać utworzony ręcznie i wskazywać na coś innego (np. na wersję
; deweloperską programu uruchamianą bezpośrednio ze źródeł).
Name: "{autodesktop}\{code:GetDesktopIconName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[Code]
// Zwraca ścieżkę, na którą wskazuje istniejący plik .lnk (przez COM
// WScript.Shell — ten sam mechanizm, którym PowerShell czyta skróty).
// Gdy odczyt się nie powiedzie — zwraca pusty tekst.
function GetShortcutTarget(LinkFile: String): String;
var
  WshShell: Variant;
  Shortcut: Variant;
begin
  Result := '';
  try
    WshShell := CreateOleObject('WScript.Shell');
    Shortcut := WshShell.CreateShortcut(LinkFile);
    Result := Shortcut.TargetPath;
  except
    Result := '';
  end;
end;

// Dobiera nazwę dla skrótu na pulpicie (bez rozszerzenia .lnk):
//   - "mBank Analyzer", jeśli pliku o tej nazwie jeszcze nie ma;
//   - ta sama nazwa, jeśli istniejący skrót wskazuje do wnętrza folderu
//     instalacji tego programu — to znaczy reinstalacja/aktualizacja
//     i można go bezpiecznie nadpisać;
//   - "mBank Analyzer (1)", "(2)", ... jeśli pod tą nazwą leży już
//     OBCY skrót (wskazujący gdzie indziej) — aby go nie nadpisać.
function GetDesktopIconName(Param: String): String;
var
  BaseName, CandidateName, LinkPath, Target, AppDir: String;
  Desktop: String;
  I: Integer;
begin
  Desktop  := ExpandConstant('{autodesktop}');
  AppDir   := ExpandConstant('{app}');
  BaseName := '{#MyAppName}';
  CandidateName := BaseName;
  I := 1;
  while True do
  begin
    LinkPath := Desktop + '\' + CandidateName + '.lnk';
    if not FileExists(LinkPath) then
      Break;

    Target := GetShortcutTarget(LinkPath);
    if (Target <> '') and (Pos(Lowercase(AppDir), Lowercase(Target)) > 0) then
      Break; // to nasz własny skrót z poprzedniej instalacji — używamy tej nazwy

    CandidateName := BaseName + ' (' + IntToStr(I) + ')';
    I := I + 1;
  end;
  Result := CandidateName;
end;

// Po usunięciu plików programu proponujemy także wyczyszczenie folderu ustawień.
// Folder %APPDATA%\mBank Analyzer zawiera config.json z danymi sprzedawcy
// i innymi ustawieniami użytkownika — deinstalator domyślnie go nie rusza,
// aby ustawienia przetrwały reinstalację/aktualizację.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    AppDataDir := ExpandConstant('{userappdata}\mBank Analyzer');
    if DirExists(AppDataDir) then
    begin
      if MsgBox(
        'Czy usunąć również folder z ustawieniami programu?' + #13#10 +
        '(dane sprzedawcy, konfiguracja)' + #13#10#13#10 +
        'Wybierz "Nie", aby zachować ustawienia — przydadzą się' + #13#10 +
        'przy ewentualnej ponownej instalacji.',
        mbConfirmation, MB_YESNO) = IDYES then
      begin
        DelTree(AppDataDir, True, True, True);
      end;
    end;
  end;
end;
