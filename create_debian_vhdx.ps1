# Requires: PowerShell 5.1+ and winget
# Run as Administrator recommended

$WorkingDir = Join-Path $HOME "Debian"
$ImageUrl = "https://cloud.debian.org/images/cloud/bookworm/latest/debian-12-genericcloud-amd64.qcow2"

$QcowFile = Join-Path $WorkingDir "debian-12-genericcloud-amd64.qcow2"
$VhdxFile = Join-Path $WorkingDir "debian-12-genericcloud-amd64.vhdx"

# Create working directory
if (-not (Test-Path $WorkingDir)) {
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
}

Write-Host "Checking for qemu-img..."

$qemuImg = Get-Command qemu-img.exe -ErrorAction SilentlyContinue

if (-not $qemuImg) {
    Write-Host "QEMU not found. Installing via winget..."

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue

    if (-not $winget) {
        throw "winget is not installed or not available."
    }

    winget install `
        --id SoftwareFreedomConservancy.QEMU `
        --accept-source-agreements `
        --accept-package-agreements `
        --silent

    # Refresh PATH from machine and user environment
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")

    $qemuImg = Get-Command qemu-img.exe -ErrorAction SilentlyContinue

    if (-not $qemuImg) {
        $qemuInstall = Get-ChildItem "C:\Program Files" -Directory -Filter "qemu*" -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($qemuInstall) {
            $candidate = Join-Path $qemuInstall.FullName "qemu-img.exe"

            if (Test-Path $candidate) {
                $qemuImg = Get-Item $candidate
            }
        }
    }

    if (-not $qemuImg) {
        throw "QEMU was installed but qemu-img.exe could not be located."
    }
}

Write-Host "Downloading Debian cloud image..."

Invoke-WebRequest `
    -Uri $ImageUrl `
    -OutFile $QcowFile

Write-Host "Converting QCOW2 to VHDX..."

$qemuExe = if ($qemuImg.Source) { $qemuImg.Source } else { $qemuImg.FullName }

& $qemuExe convert `
    -p `
    -f qcow2 `
    -O vhdx `
    $QcowFile `
    $VhdxFile

if ($LASTEXITCODE -ne 0) {
    throw "qemu-img conversion failed."
}

Write-Host ""
Write-Host "Conversion complete."
Write-Host "VHDX created at:"
Write-Host $VhdxFile -ForegroundColor Green