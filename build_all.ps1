# build_all.ps1 — полная сборка: .exe (PyInstaller) + установщик (Inno Setup)
#
# Результат: dist\mBank Analyzer Setup.exe — единственный файл, который
# нужно передать пользователю для установки программы.
#
# Требования:
#   - pip install pyinstaller pillow
#   - установленный Inno Setup 6 (winget install JRSoftware.InnoSetup)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

# ── Шаг 1: собрать автономный .exe ────────────────────────────────────────────
& "$root\build_exe.ps1"
if ($LASTEXITCODE -ne 0) { exit 1 }

# ── Шаг 2: найти компилятор Inno Setup ────────────────────────────────────────
$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if (-not $iscc) {
    $candidates = @(
        "$Env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "${Env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$Env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    $iscc = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $iscc) {
    Write-Host "Не найден компилятор Inno Setup (ISCC.exe)." -ForegroundColor Red
    Write-Host "Установите Inno Setup: winget install JRSoftware.InnoSetup" -ForegroundColor Yellow
    exit 1
}

# ── Шаг 3: скомпилировать установщик ──────────────────────────────────────────
Write-Host ""
Write-Host "==> Собираю установщик через Inno Setup..." -ForegroundColor Cyan
& "$iscc" "$root\installer.iss"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Сборка установщика завершилась с ошибкой." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host " Готово! Файл для раздачи пользователям:" -ForegroundColor Green
Write-Host "   $root\dist\mBank Analyzer Setup.exe" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
