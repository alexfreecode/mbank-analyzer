# build_exe.ps1 — buduje samodzielny plik .exe ze źródeł w src/
#
# Wynik: dist\mBank Analyzer.exe (jeden plik, --onefile)
#
# Wymagania: pip install pyinstaller pillow  (Pillow tylko do generowania ikony)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

Write-Host "==> Generuje ikone aplikacji..." -ForegroundColor Cyan
python "$root\assets\make_icon.py"

Write-Host "==> Czyszcze poprzednia kompilacje..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "$root\build" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force "$root\dist"  -ErrorAction SilentlyContinue
Remove-Item -Force "$root\*.spec"          -ErrorAction SilentlyContinue

Write-Host "==> Buduje .exe przez PyInstaller..." -ForegroundColor Cyan
pyinstaller `
    --onefile `
    --windowed `
    --name "mBank Analyzer" `
    --icon "$root\assets\app.ico" `
    --add-data "$root\assets\app.ico;assets" `
    --distpath "$root\dist" `
    --workpath "$root\build" `
    --specpath "$root" `
    --paths "$root\src" `
    "$root\src\app.py"

if ($LASTEXITCODE -ne 0) {
    Write-Host "Kompilacja zakonczyla sie bledem." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "==> Gotowe! Plik znajduje sie tutaj:" -ForegroundColor Green
Write-Host "    $root\dist\mBank Analyzer.exe"
