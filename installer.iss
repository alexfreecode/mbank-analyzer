; installer.iss — скрипт Inno Setup для сборки установщика mBank Analyzer
;
; Собрать установщик:
;   1) Сначала выполнить build_exe.ps1 — получить dist\mBank Analyzer.exe
;   2) Затем скомпилировать этот файл компилятором Inno Setup (ISCC.exe installer.iss)
;
; Результат: dist\mBank Analyzer Setup.exe — один файл для раздачи пользователям.
;
; Установка идёт в папку профиля пользователя (без прав администратора),
; создаётся ярлык в "Пуск" и (по желанию) на рабочем столе, регистрируется
; в "Установка и удаление программ" с полноценным удалением.

#define MyAppName "mBank Analyzer"
#define MyAppVersion "1.0"
; Нейтральное название издателя — никаких личных имён в установщике
; (программа должна быть свободно передаваемой без чьих-либо личных данных)
#define MyAppPublisher "mBank Analyzer"
#define MyAppExeName "mBank Analyzer.exe"

[Setup]
AppId={{B6F2B6B0-2F0A-4C7E-9C9A-7E6B6E1A9D3A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
; Устанавливаем без прав администратора — в папку текущего пользователя,
; чтобы неподготовленный пользователь не столкнулся с запросом UAC
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
; Имя ярлыка на рабочем столе вычисляется динамически (см. [Code] ниже) —
; чтобы не затереть уже существующий ярлык с тем же именем, который мог
; быть создан вручную и указывать на что-то другое (например, на dev-версию
; программы, запускаемую напрямую из исходников).
Name: "{autodesktop}\{code:GetDesktopIconName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

[Code]
// Возвращает путь, на который указывает существующий .lnk-файл (через COM
// WScript.Shell — тот же механизм, которым PowerShell читает ярлыки).
// Если прочитать не удалось — возвращает пустую строку.
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

// Подбирает имя для ярлыка на рабочем столе (без расширения .lnk):
//   - "mBank Analyzer", если файла с таким именем ещё нет;
//   - то же имя, если существующий ярлык указывает внутрь папки установки
//     этой программы — значит, это переустановка/обновление, и его
//     безопасно перезаписать;
//   - "mBank Analyzer (1)", "(2)", ... если по этому имени уже лежит
//     ЧУЖОЙ ярлык (указывающий куда-то ещё) — чтобы его не затереть.
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
      Break; // это наш собственный ярлык от предыдущей установки — переиспользуем имя

    CandidateName := BaseName + ' (' + IntToStr(I) + ')';
    I := I + 1;
  end;
  Result := CandidateName;
end;

// После удаления файлов программы предлагаем также очистить папку с настройками.
// Папка %APPDATA%\mBank Analyzer содержит config.json с реквизитами продавца
// и другими пользовательскими настройками — деинсталлятор её не трогает
// по умолчанию, чтобы настройки сохранились при переустановке/обновлении.
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
