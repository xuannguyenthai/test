# Use Server Core for a smaller footprint
FROM mcr.microsoft.com/windows/servercore:ltsc2022

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

RUN Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe" -OutFile "odt.exe"; \
    Start-Process ./odt.exe -ArgumentList '/quiet', '/passive', '/extract:.' -Wait; \
    Start-Process ./setup.exe -ArgumentList '/configure', 'configuration.xml' -Wait; \
    Write-Host "Waiting for background processes to settle..."; \
    Start-Sleep -s 30; \
    Get-Process | Where-Object {$_.Name -like '*setup*' -or $_.Name -like '*office*'} | Stop-Process -Force -ErrorAction SilentlyContinue

# Change WORKDIR away from C:\setup before deleting it
WORKDIR /
RUN Remove-Item -Path C:\setup -Recurse -Force

# 4. Set Path for Python
RUN $env:Path += ';C:\Python312;C:\Python312\Scripts'; \
    [Environment]::SetEnvironmentVariable('Path', $env:Path, [EnvironmentVariableTarget]::Machine)

CMD ["powershell"]