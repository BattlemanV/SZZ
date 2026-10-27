; Установщик программы «Отчёт СЗЗ» (Inno Setup 6).
; Собирается скриптом build.ps1 после сборки PyInstaller.

#define AppName "Отчёт СЗЗ"
#define AppVersion "1.0"
#define ExeName "Отчет СЗЗ.exe"
#define DefaultVocw "C:\integral.ltd\VocW"

[Setup]
AppId={{8E3C9C55-2D51-4B7A-9E3B-6A1C3F2D7B41}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\Отчет СЗЗ
DefaultGroupName={#AppName}
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=Отчет_СЗЗ_установка_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#ExeName}
DisableProgramGroupPage=yes

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Ярлык на рабочем столе (на него можно перетаскивать отчёты)"; GroupDescription: "Ярлыки:"

[Files]
Source: "build\dist\Отчет СЗЗ\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; шаблон пользователь может править — при обновлении не перезаписываем
Source: "..\отчет выход.docx"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "Инструкция.html"; DestDir: "{app}"; Flags: ignoreversion
; дополнения классов опасности пользователя — при обновлении не перезаписываем
Source: "классы опасности.csv"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#ExeName}"
Name: "{group}\Инструкция"; Filename: "{app}\Инструкция.html"
Name: "{group}\Классы опасности (дополнения)"; Filename: "{app}\классы опасности.csv"
Name: "{group}\Удалить {#AppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#ExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\Инструкция.html"; Description: "Открыть инструкцию"; Flags: postinstall shellexec skipifsilent nowait
Filename: "{app}\{#ExeName}"; Description: "Запустить программу"; Flags: postinstall skipifsilent nowait unchecked

[UninstallDelete]
Type: files; Name: "{app}\settings.ini"

[Code]
var
  VocwPage: TInputDirWizardPage;

function IsVocwDir(const Dir: String): Boolean;
begin
  Result := FileExists(AddBackslash(Dir) + 'DataBase\FPR.ADT') or
            FileExists(AddBackslash(Dir) + 'FPR.ADT');
end;

function InitialVocwDir(): String;
var
  Lines: TArrayOfString;
  I: Integer;
begin
  { параметр командной строки /VOCW=..., затем прежний settings.ini }
  Result := ExpandConstant('{param:VOCW|}');
  if Result <> '' then
    Exit;
  { справочник по умолчанию подставляем, только если он действительно есть }
  if IsVocwDir('{#DefaultVocw}') then
    Result := '{#DefaultVocw}';
  if LoadStringsFromFile(ExpandConstant('{localappdata}\Programs\Отчет СЗЗ\settings.ini'), Lines) then
    for I := 0 to GetArrayLength(Lines) - 1 do
      if Pos('vocw', Lowercase(Trim(Lines[I]))) = 1 then
        Result := Trim(Copy(Lines[I], Pos('=', Lines[I]) + 1, MaxInt));
end;

procedure InitializeWizard();
begin
  VocwPage := CreateInputDirPage(wpSelectDir,
    'Справочник веществ УПРЗА «Эколог» (необязательно)',
    'Где установлен справочник веществ (VocW)?',
    'Из справочника программа берёт классы опасности веществ для таблицы 3.1 ' +
    'и сверяет с ним нормативы. Укажите каталог VocW — в нём есть папка DataBase ' +
    'с файлом FPR.ADT (обычно C:\integral.ltd\VocW).' + #13#10#13#10 +
    'Если «Эколог» на этом компьютере не установлен, оставьте поле пустым: ' +
    'нормативы возьмутся из отчётов, а классы опасности можно задать ' +
    'в файле «классы опасности.csv».',
    False, '');
  VocwPage.Add('Каталог справочника (можно оставить пустым):');
  VocwPage.Values[0] := InitialVocwDir();
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = VocwPage.ID) and (Trim(VocwPage.Values[0]) <> '') and
     not IsVocwDir(VocwPage.Values[0]) then
    Result := MsgBox('В каталоге' + #13#10 + VocwPage.Values[0] + #13#10 +
      'не найден справочник (DataBase\FPR.ADT).' + #13#10#13#10 +
      'Продолжить всё равно? Каталог можно будет поменять в окне программы, ' +
      'а без справочника программа тоже работает.',
      mbConfirmation, MB_YESNO) = IDYES;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo,
  MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo + NewLine + NewLine +
    'Справочник веществ:' + NewLine + Space;
  if Trim(VocwPage.Values[0]) = '' then
    Result := Result + 'не используется'
  else
    Result := Result + VocwPage.Values[0];
  if MemoTasksInfo <> '' then
    Result := Result + NewLine + NewLine + MemoTasksInfo;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Lines: TArrayOfString;
begin
  if CurStep = ssPostInstall then
  begin
    SetArrayLength(Lines, 2);
    Lines[0] := '[paths]';
    Lines[1] := 'vocw = ' + VocwPage.Values[0];
    SaveStringsToUTF8File(ExpandConstant('{app}\settings.ini'), Lines, False);
  end;
end;
