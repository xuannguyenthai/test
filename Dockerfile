# Use Server Core for a smaller footprint
FROM mcr.microsoft.com/windows/servercore:ltsc2025

SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

# 1. Install Chocolatey
RUN Set-ExecutionPolicy Bypass -Scope Process -Force; \
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; \
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# 2. Install Python and Visual C++ Redistributable (Required for Office/Python C-extensions)
RUN choco install python --version=3.12.0 -y; \
    choco install vcredist140 -y

# 3. Download and Run Office Deployment Tool
WORKDIR /setup
COPY configuration.xml .

# These folders are required for Excel COM Automation to function on Server Core
RUN New-Item -Path 'C:\Windows\System32\config\systemprofile\Desktop' -ItemType Directory -Force; \
    New-Item -Path 'C:\Windows\SysWOW64\config\systemprofile\Desktop' -ItemType Directory -Force

RUN Invoke-WebRequest -Uri "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19725-20126.exe" \
        -OutFile "odt.exe" -UseBasicParsing; \
    $size = (Get-Item odt.exe).Length; \
    if ($size -lt 1MB) { throw "ODT download too small: $size bytes" }; \
    Start-Process ./odt.exe -ArgumentList '/quiet /extract:.' -Wait; \
    Write-Host "Downloading Office source files..."; \
    Start-Process ./setup.exe -ArgumentList '/download configuration.xml' -Wait; \
    if (-not (Test-Path 'C:\setup\officesource')) { throw "Office source download failed" }; \
    Write-Host "Installing Office from local source..."; \
    Start-Process ./setup.exe -ArgumentList '/configure configuration.xml' -Wait; \
    $timeout = 600; $elapsed = 0; \
    while (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE') -and $elapsed -lt $timeout) { \
        Write-Host "Waiting for Office... $elapsed s"; \
        Start-Sleep -s 10; \
        $elapsed += 10; \
    }; \
    if (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { \
        throw "Office installation timed out or failed" \
    }; \
    Write-Host "Office installed successfully"

RUN & 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' /regserver; \
    Start-Sleep -s 5; \
    Write-Host "COM registration complete"

# Verify COM registration
RUN $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Excel.Application\CLSID' -ErrorAction SilentlyContinue; \
    if (-not $key) { throw "Excel COM class not registered — build failed" }; \
    Write-Host "Verified: Excel.Application COM class registered at $($key.'(default)')"

# Change WORKDIR away from C:\setup before deleting it
WORKDIR /
RUN Remove-Item -Path C:\setup -Recurse -Force

# 4. Set Path for Python
RUN $env:Path += ';C:\Python312;C:\Python312\Scripts'; \
    [Environment]::SetEnvironmentVariable('Path', $env:Path, [EnvironmentVariableTarget]::Machine)

# 5. Copy entire folder and run the script
WORKDIR /app
COPY . .

RUN pip install -r requirements.txt
RUN python C:\app\xl_pdf_watcher.py data\in --interval 5 --pdf-mode standard