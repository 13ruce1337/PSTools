# --- CONFIGURATION ---

$VMName                 = "Win11-Reference"
$ISOPath                = "C:\ISOs\Win11_Unattended.iso"

$VHDSizeGB              = 80
$CPUCount               = 2

# This is the switch you'll reconnect to LATER when you're ready
# to give the VM internet/network access.
$SwitchName             = "Default Switch"

# The VM is attached to THIS switch during provisioning so it has
# no internet or LAN access until you manually reconnect it.
$ProvisioningSwitchName = "Provisioning-Isolated"

# --- HYPER-V PATHS ---

$VMPath  = Join-Path (Get-VMHost).VirtualMachinePath $VMName
$VHDPath = Join-Path (Get-VMHost).VirtualHardDiskPath "$VMName.vhdx"

# --- VALIDATION ---

if (!(Test-Path $ISOPath)) {
    throw "Windows ISO not found: $ISOPath"
}

# --- CREATE ISOLATED PROVISIONING SWITCH IF NEEDED ---

if (-not (Get-VMSwitch -Name $ProvisioningSwitchName -ErrorAction SilentlyContinue)) {

    Write-Host "Creating isolated provisioning switch: $ProvisioningSwitchName"

    New-VMSwitch `
        -Name $ProvisioningSwitchName `
        -SwitchType Private | Out-Null
}
else {
    Write-Host "Using existing isolated provisioning switch: $ProvisioningSwitchName"
}

# --- CREATE VM DIRECTORY ---

New-Item -ItemType Directory -Path $VMPath -Force | Out-Null

# --- CREATE VHD IF NEEDED ---

if (-not (Test-Path $VHDPath)) {

    Write-Host "Creating VHDX..."

    New-VHD `
        -Path $VHDPath `
        -SizeBytes ($VHDSizeGB * 1GB) `
        -Dynamic | Out-Null
}
else {
    Write-Host "Using existing VHDX: $VHDPath"
}

# --- CREATE VM IF NEEDED ---

if (-not (Get-VM -Name $VMName -ErrorAction SilentlyContinue)) {

    Write-Host "Creating VM..."

    New-VM `
        -Name $VMName `
        -MemoryStartupBytes 2GB `
        -Generation 2 `
        -SwitchName $ProvisioningSwitchName `
        -VHDPath $VHDPath `
        -Path $VMPath | Out-Null
}
else {
    Write-Host "Using existing VM: $VMName"

    # Make sure the VM's network adapter is on the isolated switch,
    # even if the VM already existed from a previous run.
    Get-VMNetworkAdapter -VMName $VMName |
        Connect-VMNetworkAdapter -SwitchName $ProvisioningSwitchName
}

# --- VM CONFIGURATION ---

Set-VMProcessor `
    -VMName $VMName `
    -Count $CPUCount

# Disable checkpoints for FFU/reference image work

Set-VM `
    -VMName $VMName `
    -CheckpointType Disabled

Set-VM `
    -VMName $VMName `
    -AutomaticCheckpointsEnabled $false

# Dynamic memory

Set-VMMemory `
    -VMName $VMName `
    -DynamicMemoryEnabled $true `
    -MinimumBytes 2GB `
    -StartupBytes 4GB `
    -MaximumBytes 8GB

# --- TPM ---

try {

    Set-VMKeyProtector `
        -VMName $VMName `
        -NewLocalKeyProtector

    Enable-VMTPM `
        -VMName $VMName

    Write-Host "vTPM enabled"
}
catch {

    Write-Warning "Unable to enable vTPM. Continuing."
}

# --- SECURE BOOT ---

try {

    Set-VMFirmware `
        -VMName $VMName `
        -EnableSecureBoot On `
        -SecureBootTemplate MicrosoftWindows

    Write-Host "Secure Boot enabled"
}
catch {

    Write-Warning "Unable to configure Secure Boot. Continuing."
}

# --- REMOVE EXISTING DVD DRIVES ---

Get-VMDvdDrive `
    -VMName $VMName `
    -ErrorAction SilentlyContinue |
    Remove-VMDvdDrive `
    -ErrorAction SilentlyContinue

# --- ATTACH WINDOWS ISO ---

Write-Host "Attaching Windows ISO..."

Add-VMDvdDrive `
    -VMName $VMName `
    -Path $ISOPath

# --- BOOT ORDER ---

$DVDDrive = Get-VMDvdDrive `
    -VMName $VMName |
    Select-Object -First 1

Set-VMFirmware `
    -VMName $VMName `
    -FirstBootDevice $DVDDrive

# --- START VM ---

if ((Get-VM $VMName).State -ne 'Running') {

    Start-VM -Name $VMName
}

# --- OPEN CONSOLE ---

vmconnect localhost $VMName

# --- SUMMARY ---

Write-Host ""
Write-Host "================================="
Write-Host "VM READY"
Write-Host "================================="
Write-Host "Name        : $VMName"
Write-Host "VM Path     : $VMPath"
Write-Host "VHD Path    : $VHDPath"
Write-Host "Windows     : $ISOPath"
Write-Host "Network     : $ProvisioningSwitchName (isolated, no internet)"
Write-Host ""
Write-Host "When you're ready to give this VM internet access, run:"
Write-Host "  Get-VMNetworkAdapter -VMName $VMName | Connect-VMNetworkAdapter -SwitchName '$SwitchName'"
