# build_exe.ps1 — собирает автономный mBankAnalyzer.exe из исходников в src/
#
# Результат: build\mBank Analyzer\mBank Analyzer.exe (один файл, --onefile)
#
# Требования: pip install pyinstaller pillow  (Pillow только для генерации иконки)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

Write-Host "==> Генерирую иконку приложения..." -ForegroundColor Cyan
python "$root\assets\make_icon.py"

Write-Host "==> Очищаю предыдущую сборку..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$root\build" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$root\dist"  -ErrorAction SilentlyContinue
Remove-Item -Force "$root\*.spec"          -ErrorAction SilentlyContinue

Write-Host "==> Собираю .exe через PyInstaller..." -ForegroundColor Cyan
pyinstaller `
    --onefile `
    --windowed `
    --name "mBank Analyzer" `
    --icon "$root\assets\app.ico" `
    --distpath "$root\dist" `
    --workpath "$root\build" `
    --specpath "$root" `
    --paths "$root\src" `
    "$root\src\app.py"

if ($LASTEXITCODE -ne 0) {
    Write-Host "Сборка завершилась с ошибкой." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "==> Готово! Файл находится здесь:" -ForegroundColor Green
Write-Host "    $root\dist\mBank Analyzer.exe"
