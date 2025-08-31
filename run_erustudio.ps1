# EruStudio PowerShell Launcher
# Run this script with: powershell -ExecutionPolicy Bypass -File run_erustudio.ps1

param(
    [switch]$InstallDeps,
    [switch]$Help
)

if ($Help) {
    Write-Host "EruStudio PowerShell Launcher" -ForegroundColor Cyan
    Write-Host "Usage:" -ForegroundColor White
    Write-Host "  .\run_erustudio.ps1              # Launch EruStudio" -ForegroundColor Gray
    Write-Host "  .\run_erustudio.ps1 -InstallDeps # Install dependencies and launch" -ForegroundColor Gray
    Write-Host "  .\run_erustudio.ps1 -Help        # Show this help message" -ForegroundColor Gray
    exit 0
}

Write-Host "Starting EruStudio..." -ForegroundColor Green
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Python not found"
    }
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python 3.8 or higher from https://python.org" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if requirements are installed
Write-Host "Checking dependencies..." -ForegroundColor Cyan
try {
    $openpyxl = pip show openpyxl 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "openpyxl not found"
    }
    Write-Host "Dependencies are installed" -ForegroundColor Green
} catch {
    if ($InstallDeps) {
        Write-Host "Installing required packages..." -ForegroundColor Yellow
        try {
            pip install -r requirements.txt
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to install packages"
            }
            Write-Host "Dependencies installed successfully" -ForegroundColor Green
        } catch {
            Write-Host "ERROR: Failed to install required packages" -ForegroundColor Red
            Write-Host "Please check your internet connection and try again" -ForegroundColor Yellow
            Write-Host ""
            Read-Host "Press Enter to exit"
            exit 1
        }
    } else {
        Write-Host "WARNING: Some dependencies are missing" -ForegroundColor Yellow
        Write-Host "Run with -InstallDeps flag to install them automatically" -ForegroundColor Yellow
        Write-Host ""
        $response = Read-Host "Continue anyway? (y/N)"
        if ($response -notmatch "^[Yy]") {
            exit 0
        }
    }
}

# Launch the application
Write-Host "Launching EruStudio..." -ForegroundColor Green
try {
    python main.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "EruStudio encountered an error. Please check the error message above." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to exit"
    }
} catch {
    Write-Host "ERROR: Failed to launch EruStudio" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
} 