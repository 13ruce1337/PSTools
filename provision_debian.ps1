# provision-vm.ps1
# requires ADK installed
# requires cloud debian converted to vhdx placed in the correct folder

param(
    [Parameter(Mandatory=$true)]
    [string]$VMName,

    [int]   $CPUs     = 1,
    [int]   $RamGB    = 2,
    [int]   $DiskGB   = 20,
    [string]$RootPass = "changeme123"
)

# --- CONFIG -------------------------------------------------------------------
$BaseVHDX  = "$env:USERPROFILE\Images\debian-12-base.vhdx"
$SeedsDir  = "$env:USERPROFILE\Seeds"
$VMRoot    = (Get-VMHost).VirtualMachinePath
$VHDRoot   = (Get-VMHost).VirtualHardDiskPath
$Switch    = "Default Switch"
$WorkDir   = "$env:TEMP\vm-provision\$VMName"
$OscdImg   = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
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
foreach ($dir in @($WorkDir, $SeedsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# --- WRITE CLOUD-INIT CONFIGS -------------------------------------------------
Write-Host "Writing cloud-init configs..."

Set-Content -Path "$WorkDir\meta-data" -Value @"
instance-id: $VMName
local-hostname: $VMName
"@

Set-Content -Path "$WorkDir\user-data" -Value @"
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

Set-Content -Path "$WorkDir\network-config" -Value @"
version: 2
ethernets:
  eth0:
    dhcp4: true
"@

# --- BUILD SEED ISO -----------------------------------------------------------
Write-Host "Building seed ISO..."

$SeedIso = "$SeedsDir\$VMName-seed.iso"
& $OscdImg -j1 -lcidata $WorkDir $SeedIso

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: oscdimg failed with exit code $LASTEXITCODE"
    exit 1
}

# --- COPY AND RESIZE BASE IMAGE -----------------------------------------------
Write-Host "Copying base image..."

$VhdPath = "$VHDRoot\$VMName.vhdx"
Copy-Item $BaseVHDX $VhdPath

$currentSize = (Get-VHD -Path $VhdPath).Size
$targetSize  = $DiskGB * 1GB

if ($targetSize -lt $currentSize) {
    $safeDiskGB = [math]::Ceiling($currentSize / 1GB)
    Write-Host "Warning: requested ${DiskGB}GB is smaller than base image ($safeDiskGB GB) -- using $safeDiskGB GB instead"
    $targetSize = $safeDiskGB * 1GB
    $DiskGB     = $safeDiskGB
}

if ($targetSize -gt $currentSize) {
    Write-Host "Resizing disk from $([math]::Round($currentSize/1GB, 1))GB to ${DiskGB}GB..."
    Resize-VHD -Path $VhdPath -SizeBytes $targetSize
} else {
    Write-Host "Disk already at ${DiskGB}GB -- skipping resize"
}

# --- CREATE VM ----------------------------------------------------------------
Write-Host "Creating VM: $VMName..."

New-VM -Name $VMName -Generation 2 `
    -MemoryStartupBytes ($RamGB * 1GB) `
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

$vm = Get-VM -Name $VMName

# --- CLEANUP ------------------------------------------------------------------
Remove-Item -Recurse -Force $WorkDir

Write-Host ""
Write-Host "Done."
Write-Host "  VM:       $VMName"
Write-Host "  CPU:      $CPUs"
Write-Host "  RAM:      ${RamGB}GB"
Write-Host "  Disk:     ${DiskGB}GB"
Write-Host "  Password: $RootPass"
Write-Host "  Connect:  ssh root@<vm-ip>"
Write-Host ""
$vm | Format-Table Name, State, CPUUsage, MemoryAssigned, Uptime, Status, Version