# Сборка установщика «Отчёт СЗЗ»: PyInstaller (программа + Python) -> Inno Setup.
# Запуск: powershell -ExecutionPolicy Bypass -File installer\build.ps1
# Результат: installer\output\Отчет_СЗЗ_установка_<версия>.exe

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$build = Join-Path $here 'build'

python -m PyInstaller --noconfirm --clean --windowed `
    --name 'Отчет СЗЗ' `
    --distpath "$build\dist" --workpath "$build\work" --specpath $build `
    --paths $root `
    --add-data "$root\data;data" `
    "$root\szz.py"
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller завершился с ошибкой' }

$iscc = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw 'Не найден Inno Setup 6 (ISCC.exe)' }

& $iscc "$here\szz.iss"
if ($LASTEXITCODE -ne 0) { throw 'Inno Setup завершился с ошибкой' }
