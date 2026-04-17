# CrowdStrike Manager - Simple Distribution Creator
# Run this to create a distributable ZIP file

$ErrorActionPreference = "Stop"

Write-Host "=== CrowdStrike Manager Distribution Creator ===" -ForegroundColor Cyan

$ProjectDir = Split-Path -Parent $PSScriptRoot
$DistDir = Join-Path $ProjectDir "dist"
$OutputDir = Join-Path $ProjectDir "Distribution"
$ZipPath = Join-Path $ProjectDir "CrowdStrikeManager.zip"

# Create output directory
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

Write-Host "`nCopying files..."
Copy-Item -Path "$DistDir\*" -Destination $OutputDir -Recurse -Force

# Create a simple batch file to run as admin
$batchContent = @'
@echo off
echo CrowdStrike Manager
echo ===================
echo.
echo Starting as Administrator...
powershell.exe -Command "Start-Process '%~dp0CrowdStrikeManager.exe' -Verb RunAs"
exit
'@

Set-Content -Path (Join-Path $OutputDir "Run as Admin.bat") -Value $batchContent

# Create README
$readme = @'
CROWDSTRIKE MANAGER
====================

REQUIREMENTS:
- Windows 10/11
- .NET 8.0 Runtime (usually pre-installed)
- Administrator privileges

INSTALLATION:
1. Copy this folder to target machine
2. Right-click "Run as Admin.bat" and select "Run as administrator"
3. Or right-click CrowdStrikeManager.exe -> "Run as administrator"

CSV FORMAT:
IPAddress,Username,Password
192.168.1.10,admin,password
192.168.1.11,DOMAIN\admin,password

CS FOLDER CONTENTS:
- certificate.cer
- certificate.pfx
- WindowsSensor.exe
- falcon_install.ps1 (optional)

OUTPUT:
Reports saved to C:\CS_Report_[timestamp]\
'@

Set-Content -Path (Join-Path $OutputDir "README.txt") -Value $readme

Write-Host "Creating ZIP archive..."
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

Compress-Archive -Path "$OutputDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal

Write-Host "`n=== Distribution Ready ===" -ForegroundColor Green
Write-Host "ZIP Location: $ZipPath"
Write-Host "Folder Location: $OutputDir"
Write-Host ""
Write-Host "Send '$ZipPath' to client" -ForegroundColor Cyan
