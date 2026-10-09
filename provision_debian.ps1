# provision_debian.ps1
# requires ADK installed
# requires cloud debian converted to vhdx placed in the correct folder

param(
    [string]$VMName   = "Debian-Server",

    [Parameter(Mandatory=$true)]
    [string]$RootPass,

    [int]   $CPUs     = 1,
    [int]   $RamGB    = 2,
    [int]   $DiskGB   = 20
)

# --- CONFIG -------------------------------------------------------------------
$BytesPerGB = [int64]1024 * 1024 * 1024
$BaseVHDX   = Join-Path $env:USERPROFILE "Images\debian-12-base.vhdx"
$TempRoot   = Join-Path $env:TEMP "vm-provision"
$WorkDir    = Join-Path $TempRoot $VMName
$SeedIso    = Join-Path $TempRoot "$VMName-seed.iso"
$VMRoot     = (Get-VMHost).VirtualMachinePath
$VHDRoot    = (Get-VMHost).VirtualHardDiskPath
$Switch     = "Default Switch"
$OscdImg    = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
# ------------------------------------------------------------------------------

# --- GUARD --------------------------------------------------------------------
if (Get-VM -Name $VMName -ErrorAction SilentlyContinue) {
    Write-Host "VM '$VMName' already exists. Exiting."
    exit 0
}

# --- VALIDATE TOOLS -----------------------------------------------------------
if (-not (Test-Path $BaseVHDX)) {
    Write-Host "ERROR: Base image not found at $BaseVHDX"
    Write-Host "Run setup-base-image.sh on your Ubuntu box and scp the result to $BaseVHDX"
    exit 1
}

if (-not (Test-Path $OscdImg)) {
    Write-Host "ERROR: oscdimg.exe not found. Is the Windows ADK installed?"
    Write-Host ""
    Write-Host "Expected path:"
    Write-Host "  $OscdImg"
    Write-Host ""
    Write-Host "Download the ADK from:"
    Write-Host "  https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install"
    Write-Host "Only the 'Deployment Tools' feature is required."
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

runcmd:
  - sed -i 's/^#*PermitRootLogin.*/PermitRootLogin yes/' /etc/ssh/sshd_config
  - sed -i 's/^#*PasswordAuthentication.*/PasswordAuthentication yes/' /etc/ssh/sshd_config
  - systemctl restart sshd
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
    Remove-Item -Recurse -Force $WorkDir -ErrorAction SilentlyContinue
    Remove-Item -Force $SeedIso -ErrorAction SilentlyContinue
    exit 1
}

# Config source files (which contain the password) are no longer needed
Remove-Item -Recurse -Force $WorkDir

# --- COPY AND RESIZE BASE IMAGE -----------------------------------------------
Write-Host "Copying base image..."

$VhdPath = Join-Path $VHDRoot "$VMName.vhdx"
Copy-Item $BaseVHDX $VhdPath

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
    Resize-VHD -Path $VhdPath -SizeBytes $targetSize
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
