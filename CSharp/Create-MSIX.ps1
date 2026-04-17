# CrowdStrike Manager MSIX Package Creator
# Run as Administrator

$ErrorActionPreference = "Stop"

Write-Host "=== CrowdStrike Manager MSIX Package Creator ===" -ForegroundColor Cyan

# Paths
$ProjectDir = Split-Path -Parent $PSScriptRoot
$PackageDir = Join-Path $ProjectDir "Package"
$OutputDir = Join-Path $ProjectDir "MSIX"
$AppDir = Join-Path $PackageDir "App"
$AssetsDir = Join-Path $AppDir "Assets"

# Create directories
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $AssetsDir -Force | Out-Null

Write-Host "`nStep 1: Copying application files..." -ForegroundColor Yellow

# Find EXE from publish folder
$PublishDir = Join-Path $ProjectDir "dist"
if (-not (Test-Path $PublishDir)) {
    $PublishDir = Join-Path $ProjectDir "CSharp\dist"
}

# Copy all files from dist folder
Copy-Item -Path "$PublishDir\*" -Destination $AppDir -Recurse -Force

# Copy manifest
Copy-Item -Path (Join-Path $PackageDir "AppxManifest.xml") -Destination $AppDir -Force

Write-Host "Step 2: Creating placeholder assets..." -ForegroundColor Yellow

# Create simple placeholder images (1x1 transparent PNG as base64)
$transparentPng = [Convert]::FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

# Generate placeholder images
$assets = @{
    "StoreLogo.png" = 50
    "Square44x44Logo.png" = 44
    "Square150x150Logo.png" = 150
    "Wide310x150Logo.png" = 310
}

foreach ($asset in $assets.GetEnumerator()) {
    $size = $asset.Value
    $path = Join-Path $AssetsDir $asset.Key
    
    # Create a simple colored square image using .NET
    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap($size, $size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::FromArgb(0, 120, 215))  # Blue background
    $graphics.Dispose()
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    
    Write-Host "  Created: $($asset.Key)"
}

Write-Host "`nStep 3: Creating self-signed certificate..." -ForegroundColor Yellow

# Create self-signed certificate
$certPassword = ConvertTo-SecureString -String "CrowdStrike2024!" -AsPlainText -Force
$cert = New-SelfSignedCertificate -Type Custom -Subject "CN=CrowdStrike Manager, O=CrowdStrike Manager, C=US" `
    -KeyUsage DigitalSignature -FriendlyName "CrowdStrike Manager Code Signing" `
    -CertStoreLocation Cert:\CurrentUser\My `
    -NotAfter (Get-Date).AddYears(5)

Write-Host "  Certificate created: $($cert.Thumbprint)"

# Export certificate with private key
$pfxPath = Join-Path $OutputDir "CrowdStrikeManager.pfx"
$certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $certPassword)
[System.IO.File]::WriteAllBytes($pfxPath, $certBytes)
Write-Host "  Exported to: $pfxPath"

Write-Host "`nStep 4: Creating MSIX package..." -ForegroundColor Yellow

# Find makeappx.exe
$makeappx = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\makeappx.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
$signtool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1

if (-not $makeappx) {
    Write-Host "  ERROR: Windows SDK not found. Please install Windows 10 SDK." -ForegroundColor Red
    Write-Host "  Download from: https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/" -ForegroundColor Yellow
    Write-Host "`nAlternative: Copy dist folder manually to client machines." -ForegroundColor Cyan
    Write-Host "  Or: Create a ZIP archive for distribution." -ForegroundColor Cyan
    exit 1
}

$makeappxPath = $makeappx.FullName
$signtoolPath = $signtool.FullName

$msixPath = Join-Path $OutputDir "CrowdStrikeManager.msix"

# Create MSIX package
& $makeappxPath pack /d $AppDir /p $msixPath /o

if ($LASTEXITCODE -eq 0) {
    Write-Host "  MSIX package created: $msixPath"
    
    Write-Host "`nStep 5: Signing MSIX package..." -ForegroundColor Yellow
    
    # Sign the package
    & $signtoolPath sign /fd SHA256 /a /f $pfxPath /p "CrowdStrike2024!" $msixPath
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Package signed successfully!" -ForegroundColor Green
    } else {
        Write-Host "  Warning: Signing failed. Package created but unsigned." -ForegroundColor Yellow
    }
} else {
    Write-Host "  ERROR: Failed to create MSIX package" -ForegroundColor Red
    exit 1
}

Write-Host "`n=== Package Created Successfully ===" -ForegroundColor Green
Write-Host "Output: $msixPath"
Write-Host "Certificate: $pfxPath"
Write-Host "Certificate Password: CrowdStrike2024!"
Write-Host ""
Write-Host "To install on client machine:" -ForegroundColor Cyan
Write-Host "1. Install certificate: Right-click $pfxPath -> Install -> Current User -> Trusted Publishers" -ForegroundColor Cyan
Write-Host "2. Double-click MSIX file to install" -ForegroundColor Cyan
