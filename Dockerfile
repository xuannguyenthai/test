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

RUN Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe" -OutFile "odt.exe"; \
    # Extract ODT
    Start-Process ./odt.exe -ArgumentList '/quiet', '/passive', '/extract:.' -Wait; \
    # Run the actual Office Installation
    Write-Host "Installing Office... this will take a while."; \
    Start-Process ./setup.exe -ArgumentList '/configure', 'configuration.xml' -Wait; \
    # Cleanup
    Remove-Item -Path C:\setup -Recycle -Force

# 4. Set Path for Python
RUN $env:Path += ';C:\Python312;C:\Python312\Scripts'; \
    [Environment]::SetEnvironmentVariable('Path', $env:Path, [EnvironmentVariableTarget]::Machine)

CMD ["powershell"]