<#
.SYNOPSIS
    Docker-optimized Excel Installer for Windows Server Core.
#>

param(
    [string]$OdtDir = "C:\ODT"
)

$ErrorActionPreference = "Stop"

Write-Host "=== Starting Docker-Optimized Excel Installation ===" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Setup Environment
# ---------------------------------------------------------------------------
if (-not (Test-Path $OdtDir)) {
    New-Item -ItemType Directory -Path $OdtDir -Force | Out-Null
}

# Ensure the Office AppData folder exists (prevents common 1603 errors in Docker)
$appDataPath = "C:\Users\ContainerAdministrator\AppData\Local\Microsoft\Office"
if (-not (Test-Path $appDataPath)) {
    New-Item -Path $appDataPath -ItemType Directory -Force | Out-Null
}

# ---------------------------------------------------------------------------
# 2. Download ODT (Using .NET WebClient for stability in Docker)
# ---------------------------------------------------------------------------
Write-Host "[1/3] Downloading Office Deployment Tool..." -ForegroundColor Yellow
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"
$odtExe = Join-Path $OdtDir "odt-setup.exe"

try {
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($odtUrl, $odtExe)
} catch {
    Write-Host "FAILED: Could not download ODT. Check internet connection." -ForegroundColor Red
    exit 1
}

Write-Host "  Extracting ODT..."
# Calling the exe directly with & is more reliable in Docker than Start-Process -Wait
& $odtExe /extract:$OdtDir /quiet
Start-Sleep -Seconds 3 # Brief pause to allow extraction to handle file locks

# ---------------------------------------------------------------------------
# 3. Create configuration XML
# ---------------------------------------------------------------------------
Write-Host "[2/3] Creating configuration..." -ForegroundColor Yellow

$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="FALSE" />
</Configuration>
"@

$configPath = Join-Path $OdtDir "excel-only.xml"
$configXml | Out-File -FilePath $configPath -Encoding UTF8
Write-Host "  Config written to: $configPath"

# ---------------------------------------------------------------------------
# 4. Run Office install
# ---------------------------------------------------------------------------
Write-Host "[3/3] Installing Excel (this takes time, please wait)..." -ForegroundColor Yellow

$setupExe = Join-Path $OdtDir "setup.exe"

# We use -PassThru to capture the ExitCode accurately
$proc = Start-Process -FilePath $setupExe -ArgumentList "/configure", "`"$configPath`"" -Wait -PassThru -NoNewWindow

# 0 = Success
# 3010 = Success (Reboot required - common and safe to ignore in Docker)
if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
    Write-Host "=== Excel installed successfully (Exit Code: $($proc.ExitCode)) ===" -ForegroundColor Green
    
    # Clean up installer files to keep Docker image small
    Remove-Item $odtExe -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "=== ERROR: Excel install failed (Exit Code: $($proc.ExitCode)) ===" -ForegroundColor Red
    Write-Host "Check ODT logs in C:\Windows\Temp (usually named like YOURCOMPUTER-*.log)"
    exit 1
}