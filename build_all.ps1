# build_all.ps1 — pelna kompilacja: .exe (PyInstaller) + instalator (Inno Setup)
#
# Wynik: dist\Suma Wplat Setup.exe — jedyny plik, ktory trzeba
# przekazac uzytkownikowi do instalacji programu.
#
# Wymagania:
#   - pip install pyinstaller pillow
#   - zainstalowany Inno Setup 6 (winget install JRSoftware.InnoSetup)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

# ── Krok 1: zbudowac samodzielny .exe ─────────────────────────────────────────
& "$root\build_exe.ps1"
if ($LASTEXITCODE -ne 0) { exit 1 }

# ── Krok 2: znalezc kompilator Inno Setup ─────────────────────────────────────
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
    Write-Host "Nie znaleziono kompilatora Inno Setup (ISCC.exe)." -ForegroundColor Red
    Write-Host "Zainstaluj Inno Setup: winget install JRSoftware.InnoSetup" -ForegroundColor Yellow
    exit 1
}

# ── Krok 3: skompilowac instalator ────────────────────────────────────────────
Write-Host ""
Write-Host "==> Buduje instalator przez Inno Setup..." -ForegroundColor Cyan
& "$iscc" "$root\installer.iss"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Kompilacja instalatora zakonczyla sie bledem." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host " Gotowe! Plik do przekazania uzytkownikom:" -ForegroundColor Green
Write-Host "   $root\dist\Suma Wplat Setup.exe" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
