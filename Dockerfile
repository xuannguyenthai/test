# Use Server Core for a smaller footprint
FROM mcr.microsoft.com/windows/servercore:ltsc2022

SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

# 1. Install Chocolatey
RUN Set-ExecutionPolicy Bypass -Scope Process -Force; \
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; \
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# 2. Install Python and Visual C++ Redistributable
RUN choco install python --version=3.12.0 -y; \
    choco install vcredist140 -y

# 3. Create Desktop directories required for Excel COM automation on Server Core
RUN New-Item -Path 'C:\Windows\System32\config\systemprofile\Desktop' -ItemType Directory -Force; \
    New-Item -Path 'C:\Windows\SysWOW64\config\systemprofile\Desktop' -ItemType Directory -Force

# 4. Enable WER crash dumps so child-process crashes are captured
RUN New-Item -Path 'C:\CrashDumps' -ItemType Directory -Force; \
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps' -Name 'DumpFolder' -Value 'C:\CrashDumps' -Type ExpandString -Force -ErrorAction SilentlyContinue; \
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps' -Name 'DumpType' -Value 2 -Type DWord -Force -ErrorAction SilentlyContinue; \
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting' -Name 'ForceQueue' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue; \
    Write-Host 'WER crash dump capture enabled'

# 5. Download Office Deployment Tool, fetch source files, and install Office
WORKDIR /setup
COPY configuration.xml .

RUN New-Item -Path 'C:\setup\logs' -ItemType Directory -Force

RUN Write-Host 'Downloading Office Deployment Tool...'; \
    Invoke-WebRequest \
        -Uri 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20158.exe' \
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
    $p = Start-Process -FilePath './setup.exe' -ArgumentList '/configure configuration.xml' -Wait -PassThru; \
    $exitCode = $p.ExitCode; \
    if ($exitCode -ne 0) { \
        Write-Host ''; \
        Write-Host '========== DIAGNOSTIC DUMP =========='; \
        Write-Host '--- Application Event Log (last 20 errors) ---'; \
        Get-EventLog -LogName Application -EntryType Error -Newest 20 -ErrorAction SilentlyContinue | \
            Format-List TimeGenerated,Source,EventID,Message | Out-String | Write-Host; \
        Write-Host '--- WER Crash Dumps ---'; \
        Get-ChildItem 'C:\CrashDumps' -ErrorAction SilentlyContinue | \
            ForEach-Object { Write-Host ('Dump: ' + $_.Name + ' (' + $_.Length + ' bytes)') }; \
        Write-Host '--- All logs under C:\Windows\Temp (recursive) ---'; \
        Get-ChildItem 'C:\Windows\Temp' -Filter '*.log' -Recurse -ErrorAction SilentlyContinue | \
            Sort-Object LastWriteTime -Descending | Select-Object -First 5 | \
            ForEach-Object { \
                Write-Host (''); \
                Write-Host ('=== ' + $_.FullName + ' ==='); \
                Get-Content $_.FullName -ErrorAction SilentlyContinue | Select-Object -Last 60 | Write-Host; \
            }; \
        Write-Host '--- All logs under C:\setup\logs (recursive) ---'; \
        Get-ChildItem 'C:\setup\logs' -Filter '*.log' -Recurse -ErrorAction SilentlyContinue | \
            Sort-Object LastWriteTime -Descending | Select-Object -First 5 | \
            ForEach-Object { \
                Write-Host (''); \
                Write-Host ('=== ' + $_.FullName + ' ==='); \
                Get-Content $_.FullName -ErrorAction SilentlyContinue | Select-Object -Last 60 | Write-Host; \
            }; \
        Write-Host '--- Office C2R logs (any location) ---'; \
        Get-ChildItem 'C:\' -Filter '*.log' -Recurse -Depth 6 -ErrorAction SilentlyContinue | \
            Where-Object { $_.Name -match 'C2R|Office|ODT|Click' } | \
            Sort-Object LastWriteTime -Descending | Select-Object -First 3 | \
            ForEach-Object { \
                Write-Host (''); \
                Write-Host ('=== ' + $_.FullName + ' ==='); \
                Get-Content $_.FullName -ErrorAction SilentlyContinue | Select-Object -Last 60 | Write-Host; \
            }; \
        Write-Host '====================================='; \
        throw ('Office installation failed: exit code ' + $exitCode); \
    }; \
    if (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { \
        throw 'EXCEL.EXE not found after installation - install may have silently failed'; \
    }; \
    Write-Host 'Office installed successfully'

# 6. Register Excel COM server and verify
RUN Write-Host 'Registering Excel COM server...'; \
    & 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' /regserver; \
    Start-Sleep -Seconds 5; \
    $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Excel.Application\CLSID' -ErrorAction SilentlyContinue; \
    if (-not $key) { throw 'Excel COM class not registered - /regserver may have failed' }; \
    Write-Host ('Verified: Excel.Application COM class registered at ' + $key.'(default)')

# 7. Clean up Office setup files to reduce image size
WORKDIR /
RUN Remove-Item -Path C:\setup -Recurse -Force; \
    Remove-Item -Path C:\CrashDumps -Recurse -Force -ErrorAction SilentlyContinue

# 8. Set PATH for Python persistently
RUN $current = [Environment]::GetEnvironmentVariable('Path', [EnvironmentVariableTarget]::Machine); \
    $additions = 'C:\Python312;C:\Python312\Scripts'; \
    if ($current -notlike '*Python312*') { \
        [Environment]::SetEnvironmentVariable('Path', ($current + ';' + $additions), [EnvironmentVariableTarget]::Machine); \
    }; \
    Write-Host 'Python PATH configured'

# 9. Copy application and install dependencies
WORKDIR /app
COPY . .

RUN pip install -r requirements.txt

RUN python C:\app\xl_pdf_watcher.py data\in --interval 5 --pdf-mode standard