$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  KindleSnapOCR" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Python check
Write-Host "[1/3] Python check..." -ForegroundColor Yellow
$pythonExists = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonExists) {
    Write-Host "  Error: Python not found" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python:" -ForegroundColor White
    Write-Host "https://www.python.org/downloads/" -ForegroundColor Cyan
    Start-Process "https://www.python.org/downloads/"
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "  OK" -ForegroundColor Green

# Virtual environment
if (-not (Test-Path "venv\Scripts\python.exe")) {
    Write-Host ""
    Write-Host "[2/3] First time setup..." -ForegroundColor Yellow
    Write-Host "  Creating virtual environment..." -ForegroundColor White
    python -m venv venv

    Write-Host "  Installing packages..." -ForegroundColor White
    & "venv\Scripts\Activate.ps1"
    pip install -q -r requirements.txt

    Write-Host "  Setup complete!" -ForegroundColor Green
} else {
    Write-Host "[2/3] Loading environment..." -ForegroundColor Yellow
    & "venv\Scripts\Activate.ps1"
    Write-Host "  Updating packages..." -ForegroundColor White
    pip install -q -r requirements.txt
    Write-Host "  OK" -ForegroundColor Green
}

Write-Host ""
Write-Host "[3/3] Starting application..." -ForegroundColor Yellow
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  GUI window will open" -ForegroundColor White
Write-Host "  You can close this window" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Start-Process pythonw -ArgumentList "main.py"
Start-Sleep -Seconds 2

Write-Host "Done! You can close this window." -ForegroundColor Green
Write-Host ""
