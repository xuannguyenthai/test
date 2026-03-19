#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs Microsoft 365 Apps (Excel only) on a GCE Windows Server VM
    using the Office Deployment Tool (ODT).

.DESCRIPTION
    Downloads ODT, creates a minimal configuration that installs only Excel,
    and runs the silent install. Uses a shared/device license so no user
    sign-in is required for COM automation.

    NOTE: You must have valid Microsoft 365 licensing (e.g. Microsoft 365 Apps
    for enterprise via volume licensing, or an E3/E5 subscription assigned to
    the VM's service account).
#>

param(
    [string]$OdtDir = "C:\ODT"
)

$ErrorActionPreference = "Stop"

Write-Host "=== Installing Excel via Office Deployment Tool ===" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Download ODT
# ---------------------------------------------------------------------------
Write-Host "`n[1/3] Downloading Office Deployment Tool ..." -ForegroundColor Yellow

if (-not (Test-Path $OdtDir)) {
    New-Item -ItemType Directory -Path $OdtDir -Force | Out-Null
}

$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe"
$odtExe = "$OdtDir\odt-setup.exe"

if (-not (Test-Path "$OdtDir\setup.exe")) {
    Write-Host "  Downloading ODT..."
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtExe -UseBasicParsing
    Write-Host "  Extracting ODT..."
    Start-Process -Wait -FilePath $odtExe -ArgumentList "/extract:$OdtDir", "/quiet"
    Remove-Item $odtExe -Force -ErrorAction SilentlyContinue
    Write-Host "  ODT extracted to: $OdtDir"
} else {
    Write-Host "  ODT already present at: $OdtDir"
}

# ---------------------------------------------------------------------------
# 2. Create configuration XML (Excel only, shared activation)
# ---------------------------------------------------------------------------
Write-Host "`n[2/3] Creating configuration ..." -ForegroundColor Yellow

$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <!-- Install ONLY Excel — exclude everything else -->
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

$configPath = "$OdtDir\excel-only.xml"
$configXml | Out-File -FilePath $configPath -Encoding UTF8
Write-Host "  Config written: $configPath"

# ---------------------------------------------------------------------------
# 3. Run Office install
# ---------------------------------------------------------------------------
Write-Host "`n[3/3] Installing Excel (this may take 5-10 minutes) ..." -ForegroundColor Yellow

$setupExe = "$OdtDir\setup.exe"
Write-Host "  Running: setup.exe /configure excel-only.xml"

$proc = Start-Process -Wait -PassThru -FilePath $setupExe `
    -ArgumentList "/configure", $configPath

if ($proc.ExitCode -eq 0) {
    Write-Host ""
    Write-Host "=== Excel installed successfully ===" -ForegroundColor Green

    # Verify
    $excelPath = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
    if ($excelPath) {
        Write-Host "  Excel path: $($excelPath.'(default)')"
    }

    Write-Host ""
    Write-Host "  Excel is ready for COM automation." -ForegroundColor Cyan
    Write-Host "  You can now start the lana-excel-watcher service:"
    Write-Host "    Start-Service lana-excel-watcher"
} else {
    Write-Host ""
    Write-Host "=== Excel install failed (exit code: $($proc.ExitCode)) ===" -ForegroundColor Red
    Write-Host "  Check the ODT logs at: $OdtDir"
    exit 1
}
