# --- CONFIGURATION ---

$VMName          = "Win11-Reference"
$ISOPath         = "C:\ISOs\Win11_Unattended.iso"

$VHDSizeGB       = 80
$CPUCount        = 2
$SwitchName      = "Default Switch"

# --- HYPER-V PATHS ---

$VMPath  = Join-Path (Get-VMHost).VirtualMachinePath $VMName
$VHDPath = Join-Path (Get-VMHost).VirtualHardDiskPath "$VMName.vhdx"

# --- VALIDATION ---

if (!(Test-Path $ISOPath)) {
    throw "Windows ISO not found: $ISOPath"
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
        -SwitchName $SwitchName `
        -VHDPath $VHDPath `
        -Path $VMPath | Out-Null
}
else {
    Write-Host "Using existing VM: $VMName"
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
Write-Host "Name      : $VMName"
Write-Host "VM Path   : $VMPath"
Write-Host "VHD Path  : $VHDPath"
Write-Host "Windows   : $ISOPath"