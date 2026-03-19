# Use Windows Server Core for full dependency support
FROM mcr.microsoft.com/windows/servercore:ltsc2022

# Set shell to PowerShell for better error handling
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

WORKDIR C:\odt

# Download the latest Office Deployment Tool (verified link)
ADD https://download.microsoft.com odtsetup.exe

# Extract the setup.exe from the ODT package
RUN Start-Process ./odtsetup.exe -ArgumentList '/quiet', '/extract:C:\odt' -Wait

# Copy your local configuration.xml into the image
COPY configuration.xml C:\odt\configuration.xml

# Run the installation for Excel only
# This step may take several minutes as it downloads Excel from the Office CDN
RUN Start-Process ./setup.exe -ArgumentList '/configure', 'C:\odt\configuration.xml' -Wait

# Clean up installer files to reduce image size
RUN Remove-Item C:\odt\odtsetup.exe -Force

# Verify Excel exists in the default installation path
RUN if (!(Test-Path 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')) { throw 'Excel installation failed' }
