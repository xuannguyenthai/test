# Use Server Core for a smaller footprint
FROM mcr.microsoft.com/windows/servercore:ltsc2025

SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

# 1. Install Chocolatey
RUN Set-ExecutionPolicy Bypass -Scope Process -Force; \
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; \
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# 2. Install Python and Visual C++ Redistributables
RUN choco install python --version=3.12.0 -y; \
    choco install vcredist140 -y

# 3. Create Desktop directories required for Excel COM automation on Server Core
RUN New-Item -Path 'C:\Windows\System32\config\systemprofile\Desktop' -ItemType Directory -Force; \
    New-Item -Path 'C:\Windows\SysWOW64\config\systemprofile\Desktop' -ItemType Directory -Force

# 4. Download Office Deployment Tool, fetch source files, and install Office
WORKDIR /setup
COPY configuration.xml .

RUN New-Item -Path 'C:\setup\logs' -ItemType Directory -Force

RUN Write-Host 'Downloading Office Deployment Tool...'; \
    Invoke-WebRequest \
        -Uri 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19725-20126.exe' \
        -OutFile 'odt.exe' -UseBasicParsing; \
    $size = (Get-Item odt.exe).Length; \
    if ($size -lt 1MB) { throw ('ODT download too small: ' + $size + ' bytes') }; \
    Write-Host 'Extracting ODT...'; \
    Start-Process -FilePath './odt.exe' -ArgumentList '/quiet /extract:.' -Wait -PassThru | Out-Null; \
    Start-Sleep -Seconds 3; \
    if (-not (Test-Path 'setup.exe')) { throw 'ODT extraction failed - setup.exe not found after extraction' }; \
    Write-Host 'ODT extracted successfully'

RUN Write-Host 'Downloading Office source files (PerpetualVL2021)...'; \
    & ./setup.exe /download configuration.xml; \
    if ($LASTEXITCODE -ne 0) { throw ('Office source download failed: exit code ' + $LASTEXITCODE) }; \
    if (-not (Test-Path 'C:\setup\officesource')) { throw 'Office source directory not created' }; \
    Write-Host 'Office source files downloaded successfully'

RUN Write-Host 'Installing Office...'; \
    & ./setup.exe /configure configuration.xml; \
    if ($LASTEXITCODE -ne 0) { \
        Write-Host '--- ODT logs (C:\setup\logs) ---'; \
        Get-ChildItem 'C:\setup\logs' -Filter '*.log' -ErrorAction SilentlyContinue | \
            Sort-Object LastWriteTime -Descending | Select-Object -First 1 | \
            ForEach-Object { Get-Content $_.FullName | Select-Object -Last 80 }; \
        Write-Host '--- Windows Temp logs ---'; \
        Get-ChildItem 'C:\Windows\Temp' -Filter '*.log' -ErrorAction SilentlyContinue | \
            Where-Object { $_.Name -match 'Office|C2R|ODT|Setup' } | \
            Sort-Object LastWriteTime -Descending | Select-Object -First 2 | \
            ForEach-Object { Write-Host ('Log: ' + $_.FullName); Get-Content $_.FullName | Select-Object -Last 40 }; \
        throw ('Office installation failed: exit code ' + $LASTEXITCODE); \
    }; \
    if (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { \
        throw 'EXCEL.EXE not found after installation - install may have silently failed'; \
    }; \
    Write-Host 'Office installed successfully'

# 5. Register Excel COM server and verify
RUN Write-Host 'Registering Excel COM server...'; \
    & 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' /regserver; \
    Start-Sleep -Seconds 5; \
    $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Excel.Application\CLSID' -ErrorAction SilentlyContinue; \
    if (-not $key) { throw 'Excel COM class not registered - /regserver may have failed' }; \
    Write-Host ('Verified: Excel.Application COM class registered at ' + $key.'(default)')

# 6. Clean up Office setup files to reduce image size
WORKDIR /
RUN Remove-Item -Path C:\setup -Recurse -Force

# 7. Set PATH for Python persistently
RUN $current = [Environment]::GetEnvironmentVariable('Path', [EnvironmentVariableTarget]::Machine); \
    $additions = 'C:\Python312;C:\Python312\Scripts'; \
    if ($current -notlike '*Python312*') { \
        [Environment]::SetEnvironmentVariable('Path', ($current + ';' + $additions), [EnvironmentVariableTarget]::Machine); \
    }; \
    Write-Host 'Python PATH configured'

# 8. Copy application and install dependencies
WORKDIR /app
COPY . .

RUN pip install -r requirements.txt


RUN python C:\app\xl_pdf_watcher.py data\in --interval 5 --pdf-mode standard