#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Provisions a GCE Windows VM for lana-excel service.
    Run this script ON the VM after RDP-ing in.

.DESCRIPTION
    Installs Python, pywin32, openpyxl, and configures the
    xl_pdf_watcher as a Windows service using NSSM.

    Prerequisites:
      - Windows Server 2022 GCE VM
      - Microsoft 365 Apps (Excel) installed via Office Deployment Tool
      - This repo cloned/copied to C:\lana-excel
#>

param(
    [string]$InstallDir = "C:\Users\xuan\lana-excel",
    [string]$PythonVersion = "3.13.5",
    [string]$OutputDir = "gs://dev-processing-data/lana-conversion-output/xl_pdf",
    [int]$WatcherInterval = 5,
    [string]$PdfMode = "both"
)

$ErrorActionPreference = "Stop"

Write-Host "=== lana-excel GCE VM Setup ===" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Install Python (via winget or direct download)
# ---------------------------------------------------------------------------
Write-Host "`n[1/5] Installing Python $PythonVersion ..." -ForegroundColor Yellow

$pythonInstalled = Get-Command python -ErrorAction SilentlyContinue
if ($pythonInstalled) {
    $ver = python --version 2>&1
    Write-Host "  Python already installed: $ver"
} else {
    Write-Host "  Downloading Python installer..."
    $pyUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
    $pyInstaller = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest -Uri $pyUrl -OutFile $pyInstaller -UseBasicParsing
    Write-Host "  Running installer (silent)..."
    Start-Process -Wait -FilePath $pyInstaller -ArgumentList `
        "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1"
    Remove-Item $pyInstaller -Force
    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    Write-Host "  Python installed: $(python --version 2>&1)"
}

# ---------------------------------------------------------------------------
# 2. Install Python dependencies
# ---------------------------------------------------------------------------
Write-Host "`n[2/5] Installing Python dependencies ..." -ForegroundColor Yellow

Push-Location $InstallDir
python -m pip install --upgrade pip
python -m pip install -r requirements.txt 2>$null
python -m pip install openpyxl pywin32

# pywin32 post-install (registers COM helpers)
Write-Host "  Running pywin32 post-install..."
python -c "import win32com; print('pywin32 OK:', win32com.__file__)"
Pop-Location

# ---------------------------------------------------------------------------
# 3. Verify Excel is installed
# ---------------------------------------------------------------------------
Write-Host "`n[3/5] Verifying Excel installation ..." -ForegroundColor Yellow

$excelPath = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
if ($excelPath) {
    Write-Host "  Excel found: $($excelPath.'(default)')"
} else {
    Write-Host "  WARNING: Excel not found in registry." -ForegroundColor Red
    Write-Host "  Install Microsoft 365 Apps using the Office Deployment Tool (ODT)."
    Write-Host "  See: deploy/install-office.ps1 or the README."
    Write-Host ""
    Write-Host "  Continuing setup — you can install Excel later before starting the service."
}

# ---------------------------------------------------------------------------
# 4. Create output directory
# ---------------------------------------------------------------------------
Write-Host "`n[4/5] Creating output directory ..." -ForegroundColor Yellow



# ---------------------------------------------------------------------------
# 5. Install NSSM and register watcher as a Windows service
# ---------------------------------------------------------------------------
Write-Host "`n[5/5] Installing xl_pdf_watcher as a Windows service ..." -ForegroundColor Yellow

$ServiceName = "lana-excel-watcher"
$nssmPath = "$InstallDir\deploy\nssm.exe"

# Download NSSM if not present
if (-not (Test-Path $nssmPath)) {
    Write-Host "  Downloading NSSM (Non-Sucking Service Manager)..."
    $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
    $nssmZip = "$env:TEMP\nssm.zip"
    Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip -UseBasicParsing
    Expand-Archive -Path $nssmZip -DestinationPath "$env:TEMP\nssm" -Force
    Copy-Item "$env:TEMP\nssm\nssm-2.24\win64\nssm.exe" $nssmPath
    Remove-Item $nssmZip -Force
    Remove-Item "$env:TEMP\nssm" -Recurse -Force
    Write-Host "  NSSM installed to: $nssmPath"
}

# Remove existing service if present
$existingSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingSvc) {
    Write-Host "  Removing existing service..."
    & $nssmPath stop $ServiceName 2>$null
    & $nssmPath remove $ServiceName confirm
}

# Find python.exe path
$pythonPath = (Get-Command python).Source

Write-Host "  Registering service: $ServiceName"
& $nssmPath install $ServiceName $pythonPath `
    "$InstallDir\xl_pdf_watcher.py" $OutputDir `
    "--interval" $WatcherInterval `
    "--pdf-mode" $PdfMode

& $nssmPath set $ServiceName DisplayName "Lana Excel PDF Watcher"
& $nssmPath set $ServiceName Description "Watches for .extract flags and converts Excel files to PDF"
& $nssmPath set $ServiceName AppDirectory $InstallDir
& $nssmPath set $ServiceName AppStdout "$InstallDir\service-stdout.log"
& $nssmPath set $ServiceName AppStderr "$InstallDir\service-stderr.log"
& $nssmPath set $ServiceName AppRotateFiles 1
& $nssmPath set $ServiceName AppRotateBytes 10485760
& $nssmPath set $ServiceName Start SERVICE_AUTO_START

Write-Host ""
Write-Host "=== Setup Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Install Excel if not already done:"
Write-Host "       .\deploy\install-office.ps1"
Write-Host "  2. Start the service:"
Write-Host "       Start-Service $ServiceName"
Write-Host "  3. Check status:"
Write-Host "       Get-Service $ServiceName"
Write-Host "  4. View logs:"
Write-Host "       Get-Content $InstallDir\service-stdout.log -Tail 50"
Write-Host ""
Write-Host "  Output directory: $OutputDir"
Write-Host "  Watcher interval: ${WatcherInterval}s"
Write-Host "  PDF mode: $PdfMode"
