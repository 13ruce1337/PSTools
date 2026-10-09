# provision_debian.ps1
# Provisions a Debian 12 VM on Hyper-V using a cloud-init seed ISO.
#
# Usage:
#   .\provision_debian.ps1                        # prompts for the root password (masked)
#   .\provision_debian.ps1 MyPass                 # VM name defaults to Debian-Server
#   .\provision_debian.ps1 MyPass WebBox          # password, then VM name
#   .\provision_debian.ps1 MyPass -CPUs 2 -RamGB 4
#
# Requirements (offered for install via winget if missing):
#   - Windows ADK "Deployment Tools" (oscdimg.exe) to build the seed ISO
#   - qemu-img (QEMU) to convert the Debian qcow2 image to a fixed VHD
#
# The Debian cloud image is downloaded to a temp folder, converted to a fixed
# VHD with qemu-img, then converted to the VM's dynamic VHDX with Hyper-V's own
# Convert-VHD. The temporary files are removed when done.

#Requires -RunAsAdministrator

param(
    [Parameter(Position=0)]
    [string]$RootPass,

    [Parameter(Position=1)]
    [string]$VMName   = "Debian-Server",

    [int]   $CPUs     = 1,
    [int]   $RamGB    = 2,
    [int]   $DiskGB   = 20,

    [string]$ImageUrl = "https://cloud.debian.org/images/cloud/bookworm/latest/debian-12-generic-amd64.qcow2"
)

# --- PROMPT FOR PASSWORD IF NOT SUPPLIED --------------------------------------
if ([string]::IsNullOrEmpty($RootPass)) {
    $secure  = Read-Host "Enter root password" -AsSecureString
    $confirm = Read-Host "Confirm root password" -AsSecureString

    $bstr1 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $bstr2 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirm)
    try {
        $RootPass        = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr1)
        $RootPassConfirm = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr2)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2)
    }

    if ([string]::IsNullOrEmpty($RootPass)) {
        Write-Host "ERROR: password cannot be empty."
        exit 1
    }
    if ($RootPass -cne $RootPassConfirm) {
        Write-Host "ERROR: passwords do not match."
        exit 1
    }
}

# --- CONFIG -------------------------------------------------------------------
$BytesPerGB = [int64]1024 * 1024 * 1024
$BlockSize  = [uint32]1048576
$TempRoot   = Join-Path $env:TEMP "vm-provision"
$ImageName  = [System.IO.Path]::GetFileName(([uri]$ImageUrl).AbsolutePath)
$ImageFile  = Join-Path $TempRoot $ImageName
$TempVhd    = Join-Path $TempRoot "$VMName-temp.vhd"
$WorkDir    = Join-Path $TempRoot $VMName
$SeedIso    = Join-Path $TempRoot "$VMName-seed.iso"
$VMRoot     = (Get-VMHost).VirtualMachinePath
$VHDRoot    = (Get-VMHost).VirtualHardDiskPath
$VhdPath    = Join-Path $VHDRoot "$VMName.vhdx"
$Switch     = "Default Switch"

$ProgramFilesX86 = ${env:ProgramFiles(x86)}
$OscdImgPath = Join-Path $ProgramFilesX86 "Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
$QemuImgPath = Join-Path $env:ProgramFiles "qemu\qemu-img.exe"

$AdkWingetId      = "Microsoft.WindowsADK"
$AdkWingetOverride = "/quiet /norestart /features OptionId.DeploymentTools"
$QemuWingetId     = "SoftwareFreedomConservancy.QEMU"
# ------------------------------------------------------------------------------

# --- HELPERS ------------------------------------------------------------------
function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machine;$user"
}

function Find-Oscdimg {
    if (Test-Path $OscdImgPath) { return $OscdImgPath }
    $cmd = Get-Command "oscdimg.exe" -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Find-QemuImg {
    $cmd = Get-Command "qemu-img.exe" -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    if (Test-Path $QemuImgPath) { return $QemuImgPath }
    return $null
}

function Confirm-Install {
    param([string]$Name, [string]$Purpose)
    Write-Host ""
    Write-Host "$Name was not found. It is needed to $Purpose."
    $answer = Read-Host "Install $Name now using winget? [Y/n]"
    return ($answer -eq '' -or $answer -match '^(y|yes)$')
}

function Install-WithWinget {
    param([string]$Id, [string]$Name, [string]$Override)

    if (-not (Get-Command "winget.exe" -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: winget is not available on this machine, so $Name cannot be installed automatically."
        return $false
    }

    $wingetArgs = @(
        'install',
        '--id', $Id,
        '--exact',
        '--source', 'winget',
        '--accept-package-agreements',
        '--accept-source-agreements'
    )
    if ($Override) {
        $wingetArgs += @('--override', $Override)
    }

    Write-Host "Installing $Name (winget id: $Id)..."
    & winget.exe @wingetArgs | Out-Host
    $exitCode = $LASTEXITCODE

    Update-SessionPath

    if ($exitCode -ne 0) {
        Write-Host "Warning: winget exited with code $exitCode while installing $Name."
        return $false
    }
    return $true
}

function Remove-ProvisionArtifacts {
    param([switch]$IncludeVhd)
    Remove-Item -Recurse -Force $WorkDir   -ErrorAction SilentlyContinue
    Remove-Item -Force $SeedIso            -ErrorAction SilentlyContinue
    Remove-Item -Force $ImageFile          -ErrorAction SilentlyContinue
    Remove-Item -Force $TempVhd            -ErrorAction SilentlyContinue
    if ($IncludeVhd) {
        Remove-Item -Force $VhdPath        -ErrorAction SilentlyContinue
    }
}
# ------------------------------------------------------------------------------

# --- GUARD --------------------------------------------------------------------
if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) {
    Write-Host "VM '$VMName' already exists. Exiting."
    exit 0
}

# --- VALIDATE / INSTALL TOOLS -------------------------------------------------
$OscdImg = Find-Oscdimg
if (-not $OscdImg) {
    if (Confirm-Install -Name "Windows ADK (Deployment Tools)" -Purpose "build the cloud-init seed ISO (oscdimg.exe)") {
        Install-WithWinget -Id $AdkWingetId -Name "Windows ADK (Deployment Tools)" -Override $AdkWingetOverride | Out-Null
        $OscdImg = Find-Oscdimg
    }
}

if (-not $OscdImg) {
    Write-Host "ERROR: oscdimg.exe not found."
    Write-Host ""
    Write-Host "Expected path:"
    Write-Host "  $OscdImgPath"
    Write-Host ""
    Write-Host "Install the ADK manually from:"
    Write-Host "  https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install"
    Write-Host "Only the 'Deployment Tools' feature is required."
    exit 1
}

$QemuImg = Find-QemuImg
if (-not $QemuImg) {
    if (Confirm-Install -Name "QEMU (qemu-img)" -Purpose "convert the Debian cloud image to a VHD") {
        Install-WithWinget -Id $QemuWingetId -Name "QEMU (qemu-img)" -Override $null | Out-Null
        $QemuImg = Find-QemuImg
    }
}

if (-not $QemuImg) {
    Write-Host "ERROR: qemu-img.exe not found."
    Write-Host ""
    Write-Host "Install it manually with:"
    Write-Host "  winget install $QemuWingetId"
    Write-Host "Then run this script again."
    exit 1
}

# --- SETUP DIRS ---------------------------------------------------------------
if (-not (Test-Path $WorkDir)) {
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
}

# --- WRITE CLOUD-INIT CONFIGS -------------------------------------------------
Write-Host "Writing cloud-init configs..."

Set-Content -Path (Join-Path $WorkDir "meta-data") -Value @"
instance-id: $VMName
local-hostname: $VMName
"@

Set-Content -Path (Join-Path $WorkDir "user-data") -Value @"
#cloud-config

hostname: $VMName
fqdn: $VMName.local

chpasswd:
  list: |
    root:$RootPass
  expire: false

ssh_pwauth: true

packages:
  - curl
  - vim

package_update: true

write_files:
  - path: /etc/ssh/sshd_config.d/10-provision.conf
    permissions: '0644'
    content: |
      PermitRootLogin yes
      PasswordAuthentication yes

runcmd:
  - systemctl restart ssh
"@

Set-Content -Path (Join-Path $WorkDir "network-config") -Value @"
version: 2
ethernets:
  eth0:
    dhcp4: true
"@

# --- BUILD SEED ISO -----------------------------------------------------------
Write-Host "Building seed ISO..."

$OscdArgs = @(
    '-j1',
    '-lcidata',
    $WorkDir,
    $SeedIso
)
& $OscdImg @OscdArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: oscdimg failed with exit code $LASTEXITCODE"
    Remove-ProvisionArtifacts
    exit 1
}

# Config source files (which contain the password) are no longer needed
Remove-Item -Recurse -Force $WorkDir

# --- DOWNLOAD DEBIAN IMAGE ----------------------------------------------------
Write-Host "Downloading Debian cloud image..."
Write-Host "  $ImageUrl"

# The progress bar makes Invoke-WebRequest extremely slow in Windows PowerShell 5.1
$previousProgress   = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'
try {
    Invoke-WebRequest -Uri $ImageUrl -OutFile $ImageFile -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Host "ERROR: download failed: $($_.Exception.Message)"
    Remove-ProvisionArtifacts
    exit 1
} finally {
    $ProgressPreference = $previousProgress
}

# --- CONVERT IMAGE TO FIXED VHD (qemu-img) ------------------------------------
Write-Host "Converting image to a temporary fixed VHD..."

$QemuArgs = @(
    'convert',
    '-f', 'qcow2',
    '-O', 'vpc',
    '-o', 'subformat=fixed,force_size',
    $ImageFile,
    $TempVhd
)
& $QemuImg @QemuArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: qemu-img failed with exit code $LASTEXITCODE"
    Remove-ProvisionArtifacts
    exit 1
}

# The downloaded image is no longer needed
Remove-Item -Force $ImageFile -ErrorAction SilentlyContinue

# qemu-img on Windows may write its output as an NTFS sparse file. Hyper-V
# refuses to open sparse virtual disks, so clear the sparse attribute.
& fsutil.exe sparse setflag $TempVhd 0 | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: could not clear the sparse flag on $TempVhd (fsutil exit code $LASTEXITCODE)"
    Remove-ProvisionArtifacts
    exit 1
}

# --- CONVERT FIXED VHD TO DYNAMIC VHDX (Hyper-V) ------------------------------
Write-Host "Converting to dynamic VHDX with Hyper-V at $VhdPath..."

try {
    Convert-VHD -Path $TempVhd -DestinationPath $VhdPath -VHDType Dynamic -BlockSizeBytes $BlockSize -ErrorAction Stop
} catch {
    Write-Host "ERROR: Convert-VHD failed: $($_.Exception.Message)"
    Remove-ProvisionArtifacts -IncludeVhd
    exit 1
}

# The temporary fixed VHD is no longer needed
Remove-Item -Force $TempVhd -ErrorAction SilentlyContinue

# --- RESIZE DISK --------------------------------------------------------------
# Note: the partition table's backup GPT header will not be at the end of the
# grown disk until the guest fixes it. Debian cloud images do this on first boot
# (cloud-init growpart), so a kernel "alternate GPT header" message is expected.
$currentSize = [int64](Get-VHD -Path $VhdPath).Size
$targetSize  = [int64]$DiskGB * $BytesPerGB

if ($targetSize -lt $currentSize) {
    $safeDiskGB = [int][math]::Ceiling($currentSize / $BytesPerGB)
    Write-Host "Warning: requested ${DiskGB}GB is smaller than base image (${safeDiskGB}GB) -- using ${safeDiskGB}GB instead"
    $targetSize = [int64]$safeDiskGB * $BytesPerGB
    $DiskGB     = $safeDiskGB
}

if ($targetSize -gt $currentSize) {
    $currentGBRounded = [math]::Round($currentSize / $BytesPerGB, 1)
    Write-Host "Resizing disk from ${currentGBRounded}GB to ${DiskGB}GB..."
    try {
        Resize-VHD -Path $VhdPath -SizeBytes $targetSize -ErrorAction Stop
    } catch {
        Write-Host "ERROR: disk resize failed: $($_.Exception.Message)"
        Remove-ProvisionArtifacts -IncludeVhd
        exit 1
    }
} else {
    Write-Host "Disk already at ${DiskGB}GB -- skipping resize"
}

# --- CREATE VM ----------------------------------------------------------------
Write-Host "Creating VM: $VMName..."

$memoryBytes = [int64]$RamGB * $BytesPerGB

New-VM -Name $VMName -Generation 2 `
    -MemoryStartupBytes $memoryBytes `
    -SwitchName $Switch `
    -Path $VMRoot | Out-Null

Add-VMHardDiskDrive -VMName $VMName -Path $VhdPath
Add-VMDvdDrive      -VMName $VMName -Path $SeedIso
Set-VMProcessor     -VMName $VMName -Count $CPUs
Set-VMFirmware      -VMName $VMName -EnableSecureBoot Off

$disk = Get-VMHardDiskDrive -VMName $VMName
Set-VMFirmware -VMName $VMName -FirstBootDevice $disk

# --- START VM AND WAIT FOR RUNNING STATE --------------------------------------
Start-VM -Name $VMName

Write-Host "Waiting for VM to start..."
$timeout = 30
$elapsed = 0
while ((Get-VM -Name $VMName).State -eq 'Off') {
    if ($elapsed -ge $timeout) {
        Write-Host "Warning: VM did not start within $timeout seconds"
        break
    }
    Start-Sleep -Seconds 1
    $elapsed++
}

# --- WAIT FOR GUEST OS, THEN REMOVE SEED ISO ----------------------------------
# cloud-init reads the seed very early in first boot. Once the guest heartbeat
# is up, the seed has been consumed and the ISO is no longer needed.
Write-Host "Waiting for guest heartbeat before removing seed ISO..."
$hbTimeout = 300
$hbElapsed = 0
$hbOk      = $false
while ($hbElapsed -lt $hbTimeout) {
    $hb = Get-VMIntegrationService -VMName $VMName -Name 'Heartbeat' -ErrorAction SilentlyContinue
    if ($hb -and $hb.PrimaryStatusDescription -eq 'OK') {
        $hbOk = $true
        break
    }
    Start-Sleep -Seconds 2
    $hbElapsed += 2
}

if ($hbOk) {
    Write-Host "Guest is up. Ejecting and deleting seed ISO..."
    Get-VMDvdDrive -VMName $VMName | Set-VMDvdDrive -Path $null
    Remove-Item -Force $SeedIso -ErrorAction SilentlyContinue
    if (Test-Path $SeedIso) {
        Write-Host "Warning: could not delete seed ISO at $SeedIso"
    }
} else {
    Write-Host "Warning: no guest heartbeat within $hbTimeout seconds -- leaving seed ISO in place at:"
    Write-Host "  $SeedIso"
    Write-Host "Eject the DVD and delete the ISO manually once the VM has booted."
}

# --- CLEAN UP TEMP FOLDER IF EMPTY --------------------------------------------
if ((Test-Path $TempRoot) -and -not (Get-ChildItem -Path $TempRoot -Force -ErrorAction SilentlyContinue)) {
    Remove-Item -Force $TempRoot -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "Done."
Write-Host "  VM:       $VMName"
Write-Host "  CPU:      $CPUs"
Write-Host "  RAM:      ${RamGB}GB"
Write-Host "  Disk:     ${DiskGB}GB"
Write-Host "  Connect:  ssh root@<vm-ip>"
Write-Host ""

$vmColumns = @('Name', 'State', 'CPUUsage', 'MemoryAssigned', 'Uptime', 'Status', 'Version')
Get-VM -Name $VMName | Format-Table -AutoSize -Property $vmColumns | Out-Host
