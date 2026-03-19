WORKDIR /setup
COPY configuration.xml .

RUN New-Item -Path 'C:\Windows\System32\config\systemprofile\Desktop' -ItemType Directory -Force; \
    New-Item -Path 'C:\Windows\SysWOW64\config\systemprofile\Desktop' -ItemType Directory -Force

RUN Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkID=626065" -OutFile "odt.exe"; \
    Start-Process ./odt.exe -ArgumentList '/quiet /extract:.' -Wait; \
    Start-Process ./setup.exe -ArgumentList '/configure configuration.xml' -Wait; \
    $timeout = 600; $elapsed = 0; \
    while (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE') -and $elapsed -lt $timeout) { \
        Start-Sleep -s 10; $elapsed += 10; \
    }; \
    if (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { throw "Office install failed" }

RUN & 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' /regserver; \
    Start-Sleep -s 5

# Verify COM registration
RUN $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Classes\Excel.Application\CLSID' -ErrorAction SilentlyContinue; \
    if (-not $key) { throw "Excel COM class not registered" }; \
    Write-Host "COM registration OK"

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