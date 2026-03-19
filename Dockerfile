# escape=`
FROM mcr.microsoft.com/windows/servercore/iis:windowsservercore-ltsc2025

SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]



# ── Copy entire project folder ───────────────────────────────────────────────
COPY . C:/app/

WORKDIR C:/app

# ── Run install scripts as Administrator ─────────────────────────────────────
RUN powershell -ExecutionPolicy Bypass -File C:\app\deploy\install-office.ps1
RUN powershell -ExecutionPolicy Bypass -File C:\app\deploy\setup-gce-vm.ps1

# ── Install Python dependencies ──────────────────────────────────────────────
RUN pip install --no-cache-dir -r C:\app\requirements.txt

# ── Create a restricted local user ──────────────────────────────────────────
RUN net user appuser SecureP@ss123! /add; `
    net localgroup Users appuser /add

# ── Switch to normal user ────────────────────────────────────────────────────
USER appuser

CMD ["python", "xl_pdf_watcher.py"]