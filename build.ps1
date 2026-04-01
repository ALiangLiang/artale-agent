# Artale Helper Build Script (PowerShell)

# 1. Clean previous builds
Write-Host "Cleaning up old build artifacts..." -ForegroundColor Cyan
if (Test-Path "./dist") { Remove-Item -Path "./dist" -Recurse -Force }
if (Test-Path "./build") { Remove-Item -Path "./build" -Recurse -Force }
if (Test-Path "*.spec") { Remove-Item -Path "*.spec" -Force }

# 2. Run PyInstaller
Write-Host "Starting PyInstaller build..." -ForegroundColor Green

# Parameters:
# --onefile: Bundles everything into a single EXE
# --noconsole: Skips the black black terminal window when running
# --name: The output file name
# --add-data: Includes the buff_pngs folder inside the EXE (Syntax: "source;destination")
# --hidden-import: Ensures extra dependencies are caught
# --clean: Cleans cache before building

python -m PyInstaller `
    --onefile `
    --noconsole `
    --name "ArtaleAgent" `
    --add-data "buff_pngs;buff_pngs" `
    --hidden-import "psutil" `
    --hidden-import "pynput.keyboard._win32" `
    --clean `
    main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nSuccessfully built dist/ArtaleAgent.exe!" -ForegroundColor Green
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
}
