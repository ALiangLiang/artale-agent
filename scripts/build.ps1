# 0. Kill existing process to avoid "Access Denied" on dist folder
Write-Host "Closing existing ArtaleAgent processes..." -ForegroundColor Yellow
Stop-Process -Name "ArtaleAgent" -ErrorAction SilentlyContinue

# 1. Clean previous builds
Write-Host "Cleaning up old build artifacts..." -ForegroundColor Cyan
if (Test-Path "./dist") { Remove-Item -Path "./dist" -Recurse -Force }
if (Test-Path "./build") { Remove-Item -Path "./build" -Recurse -Force }

# 2. Run PyInstaller
Write-Host "Starting PyInstaller build..." -ForegroundColor Green

uv run build-win

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nSuccessfully built dist/ArtaleAgent.exe!" -ForegroundColor Green
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
}
