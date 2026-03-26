# build_exe.ps1
# Baut dist\excel-editor.exe als standalone Windows-Executable.
#
# Voraussetzung: Python 3.8+ auf Windows installiert (python.org)
# Ausführen in Windows PowerShell (nicht WSL):
#   cd <Projektordner>
#   .\build_exe.ps1

$ErrorActionPreference = "Stop"

Write-Host "=== excel-editor EXE Build ===" -ForegroundColor Cyan
Write-Host ""

# Zum Skript-Verzeichnis wechseln
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Python-Executable bestimmen (py-Launcher hat Vorrang auf Windows)
$pythonCmd = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
} else {
    Write-Host "FEHLER: Python nicht gefunden. Bitte von https://python.org installieren." -ForegroundColor Red
    exit 1
}
Write-Host "  Python-Befehl: $pythonCmd ($( & $pythonCmd --version 2>&1 ))" -ForegroundColor Gray
Write-Host ""

# Abhängigkeiten installieren
Write-Host "1/3  Installiere Abhängigkeiten ..." -ForegroundColor Yellow
& $pythonCmd -m pip install -e . --quiet
& $pythonCmd -m pip install pyinstaller --quiet
Write-Host "     OK" -ForegroundColor Green

# Altes Build-Verzeichnis bereinigen
if (Test-Path "dist\excel-editor.exe") {
    Write-Host "     Altes dist\excel-editor.exe wird überschrieben."
}

# EXE bauen
Write-Host "2/3  Baue EXE (kann 1–2 Minuten dauern) ..." -ForegroundColor Yellow
& $pythonCmd -m PyInstaller excel_editor.spec --clean --noconfirm
Write-Host "     OK" -ForegroundColor Green

# Ergebnis prüfen
$exePath = "dist\excel-editor.exe"
if (Test-Path $exePath) {
    $sizeMB = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
    $fullPath = (Resolve-Path $exePath).Path
    Write-Host ""
    Write-Host "3/3  Fertig!" -ForegroundColor Green
    Write-Host "     EXE:  $fullPath  ($sizeMB MB)"
    Write-Host ""
    Write-Host "Verwendung:" -ForegroundColor Cyan
    Write-Host "  .\dist\excel-editor.exe --file `"C:\Pfad\zur\Datei.xlsx`" --sheet Sheet1 --rows 10"
    Write-Host "  .\dist\excel-editor.exe --file `"C:\Pfad\zur\Datei.xlsx`" --sheet Sheet1 --move-from 1010 --move-after 1026 --output `"C:\Users\%USERNAME%\Downloads\ergebnis.xlsx`""
} else {
    Write-Host "FEHLER: EXE nicht gefunden. Siehe Build-Output oben." -ForegroundColor Red
    exit 1
}
