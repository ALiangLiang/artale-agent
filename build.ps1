# 0. Kill existing process to avoid "Access Denied" on dist folder
Write-Host "Closing existing ArtaleAgent processes..." -ForegroundColor Yellow
Stop-Process -Name "ArtaleAgent" -ErrorAction SilentlyContinue

# 1. Clean previous builds
Write-Host "Cleaning up old build artifacts..." -ForegroundColor Cyan
if (Test-Path "./dist") { Remove-Item -Path "./dist" -Recurse -Force }
if (Test-Path "./build") { Remove-Item -Path "./build" -Recurse -Force }

# 2. Run PyInstaller
Write-Host "Starting PyInstaller build..." -ForegroundColor Green

uv run python -m PyInstaller `
    --onefile `
    --name "ArtaleAgent" `
    --icon "app_icon.ico" `
    --add-data "buff_pngs;buff_pngs" `
    --add-data "Tesseract-OCR;Tesseract-OCR" `
    --add-data "app_icon.png;." `
    --add-data "VERSION;." `
    --hidden-import "psutil" `
    --hidden-import "pynput.keyboard._win32" `
    --hidden-import "win32process" `
    --hidden-import "win32file" `
    --hidden-import "PyQt6.QtCore" `
    --hidden-import "PyQt6.QtGui" `
    --hidden-import "PyQt6.QtWidgets" `
    --hidden-import "PyQt6.QtNetwork" `
    --hidden-import "PyQt6.QtWebSockets" `
    --hidden-import "sip" `
    --clean `
    --noconsole `
    --noupx `
    main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nSuccessfully built dist/ArtaleAgent.exe!" -ForegroundColor Green
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
}
