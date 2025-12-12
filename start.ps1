$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  KindleSnapOCR" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Python check
Write-Host "[1/4] Checking Python..." -ForegroundColor Yellow
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
    Write-Host "[2/4] First time setup..." -ForegroundColor Yellow
    Write-Host "  Creating virtual environment..." -ForegroundColor White
    python -m venv venv

    Write-Host "  Installing packages..." -ForegroundColor White
    & "venv\Scripts\Activate.ps1"
    pip install -q -r requirements.txt

    Write-Host "  Setup complete!" -ForegroundColor Green
} else {
    Write-Host "[2/4] Loading environment..." -ForegroundColor Yellow
    & "venv\Scripts\Activate.ps1"
    pip install -q -r requirements.txt
    Write-Host "  OK" -ForegroundColor Green
}

# Tesseract OCR check
Write-Host ""
Write-Host "[3/4] Checking Tesseract OCR..." -ForegroundColor Yellow

$tesseractPaths = @(
    "C:\Program Files\Tesseract-OCR\tesseract.exe",
    "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe"
)

$tesseractFound = $false
foreach ($path in $tesseractPaths) {
    if (Test-Path $path) {
        $tesseractFound = $true
        Write-Host "  Found: $path" -ForegroundColor Green
        break
    }
}

if (-not $tesseractFound) {
    Write-Host "  Tesseract not found" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Tesseract is required for OCR." -ForegroundColor White
    Write-Host "  Install now? (Y/N)" -ForegroundColor Cyan
    $response = Read-Host "  Select"

    if ($response -eq "Y" -or $response -eq "y") {
        Write-Host ""
        Write-Host "  Downloading and installing Tesseract..." -ForegroundColor Yellow
        Write-Host "  (Admin permission may be required)" -ForegroundColor Gray
        Write-Host ""

        python install_tesseract.py

        if ($LASTEXITCODE -eq 0) {
            Write-Host ""
            Write-Host "  Tesseract installed successfully!" -ForegroundColor Green
        } else {
            Write-Host ""
            Write-Host "  Installation failed." -ForegroundColor Red
            Write-Host "  Please install manually:" -ForegroundColor White
            Write-Host "  https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  Continuing without OCR..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Continuing without OCR..." -ForegroundColor Yellow
    }
} else {
    Write-Host "  OK" -ForegroundColor Green
}

Write-Host ""
Write-Host "[4/4] Starting application..." -ForegroundColor Yellow
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
