function New-DebianVM {
    <#
    .SYNOPSIS
        Creates and starts a Generation 2 Hyper-V VM from a Debian cloud image.

    .DESCRIPTION
        Copies a base Debian cloud VHDX, resizes it, builds a cloud-init NoCloud
        seed ISO with oscdimg (Windows ADK Deployment Tools), creates the VM,
        attaches both, and starts it.

        If the base VHDX does not exist (or -RebuildBaseImage is specified), it
        is created automatically: qemu-img is located (QEMU is installed via
        winget if needed), the Debian cloud qcow2 image is downloaded, and it is
        converted to VHDX at -BaseVhdxPath.

        Requirements:
          - Elevated session with the Hyper-V PowerShell module
          - Windows ADK "Deployment Tools" (oscdimg.exe)
          - winget (only if the base image must be built and QEMU is not
            already installed)

    .PARAMETER Name
        VM name. Also used as the guest hostname.

    .PARAMETER CpuCount
        Number of virtual processors. Default 1.

    .PARAMETER MemoryGB
        Startup memory in GB. Default 2.

    .PARAMETER DiskGB
        Target disk size in GB. If smaller than the base image, the base image
        size is used instead.

    .PARAMETER RootPassword
        Root password for the guest, as a SecureString.

    .PARAMETER BaseVhdxPath
        Path to the base Debian VHDX. Created automatically if it does not exist.

    .PARAMETER BaseImageUrl
        URL of the Debian cloud qcow2 image used to build the base VHDX.

    .PARAMETER RebuildBaseImage
        Download and convert the base image again, replacing any existing
        base VHDX.

    .PARAMETER SeedsPath
        Folder where the cloud-init seed ISO is stored.

    .PARAMETER SwitchName
        Hyper-V virtual switch to connect to.

    .PARAMETER OscdimgPath
        Path to oscdimg.exe.

    .EXAMPLE
        $pw = Read-Host -AsSecureString -Prompt 'Root password'
        New-DebianVM -Name web01 -CpuCount 2 -MemoryGB 4 -DiskGB 40 -RootPassword $pw

    .EXAMPLE
        New-DebianVM -Name test01 -RootPassword $pw -WhatIf

    .EXAMPLE
        New-DebianVM -Name test02 -RootPassword $pw -RebuildBaseImage

    .OUTPUTS
        Microsoft.HyperV.PowerShell.VirtualMachine
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType('Microsoft.HyperV.PowerShell.VirtualMachine')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9-]{0,62}$')]
        [string]$Name,

        [ValidateRange(1, 64)]
        [int]$CpuCount = 1,

        [ValidateRange(1, 1024)]
        [int]$MemoryGB = 2,

        [ValidateRange(1, 65536)]
        [int]$DiskGB = 20,

        [Parameter(Mandatory)]
        [securestring]$RootPassword,

        [ValidateNotNullOrEmpty()]
        [string]$BaseVhdxPath = (Join-Path $env:USERPROFILE 'Images\debian-12-base.vhdx'),

        [ValidateNotNullOrEmpty()]
        [string]$BaseImageUrl = 'https://cloud.debian.org/images/cloud/bookworm/latest/debian-12-genericcloud-amd64.qcow2',

        [switch]$RebuildBaseImage,

        [ValidateNotNullOrEmpty()]
        [string]$SeedsPath = (Join-Path $env:USERPROFILE 'Seeds'),

        [ValidateNotNullOrEmpty()]
        [string]$SwitchName = 'Default Switch',

        [ValidateNotNullOrEmpty()]
        [string]$OscdimgPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe'
    )

    # --- Inner helpers --------------------------------------------------------
    # Defined inline so they are not exported as part of the module.

    # Returns the full path to qemu-img.exe, installing QEMU via winget if needed.
    function Get-QemuImgPath {
        $found = Get-Command qemu-img.exe -ErrorAction SilentlyContinue
        if ($found) {
            return $found.Source
        }

        Write-Verbose 'qemu-img not found. Installing QEMU via winget...'

        if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
            throw 'qemu-img is not installed and winget is not available to install it.'
        }

        & winget.exe install `
            --id SoftwareFreedomConservancy.QEMU `
            --accept-source-agreements `
            --accept-package-agreements `
            --silent | Out-Host

        if ($LASTEXITCODE -ne 0) {
            Write-Warning "winget exited with code $LASTEXITCODE; checking whether qemu-img is available anyway."
        }

        # Refresh PATH from machine and user environment.
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                    [System.Environment]::GetEnvironmentVariable('Path', 'User')

        $found = Get-Command qemu-img.exe -ErrorAction SilentlyContinue
        if ($found) {
            return $found.Source
        }

        # Fall back to the default install location.
        $qemuInstall = Get-ChildItem -LiteralPath $env:ProgramFiles -Directory -Filter 'qemu*' -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($qemuInstall) {
            $candidate = Join-Path $qemuInstall.FullName 'qemu-img.exe'
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }

        throw 'QEMU was installed but qemu-img.exe could not be located.'
    }

    # Downloads the Debian cloud qcow2 image and converts it to VHDX.
    function New-DebianBaseVhdx {
        param(
            [Parameter(Mandatory)][string]$Url,
            [Parameter(Mandatory)][string]$Destination
        )

        $imageWorkDir = Join-Path $env:TEMP 'PSTools\New-DebianVM\_baseimage'
        $qcowFile     = Join-Path $imageWorkDir ([System.IO.Path]::GetFileName(([uri]$Url).LocalPath))
        $tempVhdx     = Join-Path $imageWorkDir 'base-conversion.vhdx'
        $destDir      = Split-Path -Path $Destination -Parent

        try {
            foreach ($dir in @($imageWorkDir, $destDir)) {
                if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                    $null = New-Item -ItemType Directory -Path $dir -Force
                }
            }

            $qemuExe = Get-QemuImgPath

            Write-Verbose "Downloading Debian cloud image from $Url"
            $oldProgress = $ProgressPreference
            try {
                # The progress bar makes Invoke-WebRequest extremely slow in Windows PowerShell 5.1.
                $ProgressPreference = 'SilentlyContinue'
                Invoke-WebRequest -Uri $Url -OutFile $qcowFile -UseBasicParsing -ErrorAction Stop
            }
            finally {
                $ProgressPreference = $oldProgress
            }

            Write-Verbose 'Converting QCOW2 to VHDX...'
            & $qemuExe convert -p -f qcow2 -O vhdx $qcowFile $tempVhdx | Out-Host
            if ($LASTEXITCODE -ne 0) {
                throw "qemu-img conversion failed with exit code $LASTEXITCODE."
            }

            # Only replace the base image once conversion has fully succeeded.
            Move-Item -LiteralPath $tempVhdx -Destination $Destination -Force -ErrorAction Stop
            Write-Verbose "Base image created at $Destination"
        }
        finally {
            if (Test-Path -LiteralPath $imageWorkDir) {
                Remove-Item -LiteralPath $imageWorkDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # --- Preflight ------------------------------------------------------------
    if (-not (Test-IsAdmin)) {
        throw 'New-DebianVM must be run from an elevated session.'
    }

    if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
        throw 'The Hyper-V PowerShell module is not installed. Enable the Hyper-V management tools first.'
    }

    if (Get-VM -Name $Name -ErrorAction SilentlyContinue) {
        Write-Warning "VM '$Name' already exists. Nothing to do."
        return
    }

    if (-not (Test-Path -LiteralPath $OscdimgPath)) {
        throw ("oscdimg.exe not found at '$OscdimgPath'. Install the Windows ADK 'Deployment Tools' feature: " +
               'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install')
    }

    $buildBaseImage = $RebuildBaseImage -or -not (Test-Path -LiteralPath $BaseVhdxPath)

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create and start Debian VM')) {
        return
    }

    # --- Base image -----------------------------------------------------------
    if ($buildBaseImage) {
        if ($RebuildBaseImage) {
            Write-Verbose "Rebuilding base image at '$BaseVhdxPath'."
        }
        else {
            Write-Verbose "Base image not found at '$BaseVhdxPath'; building it."
        }

        New-DebianBaseVhdx -Url $BaseImageUrl -Destination $BaseVhdxPath
    }

    # --- Paths ----------------------------------------------------------------
    $vmHost  = Get-VMHost
    $vmRoot  = $vmHost.VirtualMachinePath
    $vhdRoot = $vmHost.VirtualHardDiskPath
    $workDir = Join-Path $env:TEMP "PSTools\New-DebianVM\$Name"
    $seedIso = Join-Path $SeedsPath "$Name-seed.iso"
    $vhdPath = Join-Path $vhdRoot "$Name.vhdx"

    # Explicit byte multiplier (avoids relying on the 1GB numeric literal suffix).
    [int64]$bytesPerGB = 1073741824

    # Plain-text password is only needed to write the cloud-init file.
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($RootPassword)
    try {
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    try {
        foreach ($dir in @($workDir, $SeedsPath)) {
            if (-not (Test-Path -LiteralPath $dir)) {
                $null = New-Item -ItemType Directory -Path $dir -Force
            }
        }

        # --- cloud-init configs -----------------------------------------------
        Write-Verbose 'Writing cloud-init configs...'

        # cloud-init rejects a BOM and prefers LF, so write UTF-8 (no BOM) with LF.
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writeFile = {
            param([string]$Path, [string[]]$Lines)
            [System.IO.File]::WriteAllText($Path, (($Lines -join "`n") + "`n"), $utf8NoBom)
        }

        & $writeFile (Join-Path $workDir 'meta-data') @(
            "instance-id: $Name"
            "local-hostname: $Name"
        )

        & $writeFile (Join-Path $workDir 'user-data') @(
            '#cloud-config'
            ''
            "hostname: $Name"
            "fqdn: $Name.local"
            ''
            'chpasswd:'
            '  list: |'
            "    root:$plainPassword"
            '  expire: false'
            ''
            'ssh_pwauth: true'
            ''
            'packages:'
            '  - curl'
            '  - vim'
            ''
            'package_update: true'
            ''
            'runcmd:'
            "  - sed -i 's/^#*PermitRootLogin.*/PermitRootLogin yes/' /etc/ssh/sshd_config"
            "  - sed -i 's/^#*PasswordAuthentication.*/PasswordAuthentication yes/' /etc/ssh/sshd_config"
            '  - systemctl restart sshd'
        )

        & $writeFile (Join-Path $workDir 'network-config') @(
            'version: 2'
            'ethernets:'
            '  eth0:'
            '    dhcp4: true'
        )

        # --- Seed ISO ---------------------------------------------------------
        Write-Verbose "Building seed ISO: $seedIso"

        $oscdimgArgs = @('-j1', '-lcidata', $workDir, $seedIso)
        & $OscdimgPath @oscdimgArgs | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "oscdimg failed with exit code $LASTEXITCODE."
        }

        # --- Disk -------------------------------------------------------------
        Write-Verbose "Copying base image to $vhdPath"
        Copy-Item -LiteralPath $BaseVhdxPath -Destination $vhdPath -ErrorAction Stop

        $currentSize = (Get-VHD -Path $vhdPath).Size
        $targetSize  = [int64]$DiskGB * $bytesPerGB

        if ($targetSize -lt $currentSize) {
            $safeDiskGB = [int][math]::Ceiling($currentSize / $bytesPerGB)
            Write-Warning "Requested ${DiskGB}GB is smaller than the base image; using ${safeDiskGB}GB instead."
            $DiskGB     = $safeDiskGB
            $targetSize = [int64]$DiskGB * $bytesPerGB
        }

        if ($targetSize -gt $currentSize) {
            $currentSizeGB = [math]::Round($currentSize / $bytesPerGB, 1)
            Write-Verbose ('Resizing disk from {0}GB to {1}GB...' -f $currentSizeGB, $DiskGB)
            Resize-VHD -Path $vhdPath -SizeBytes $targetSize
        }
        else {
            Write-Verbose "Disk already at ${DiskGB}GB; skipping resize."
        }

        # --- VM ---------------------------------------------------------------
        Write-Verbose "Creating VM: $Name"

        $memoryBytes = [int64]$MemoryGB * $bytesPerGB

        $null = New-VM -Name $Name -Generation 2 `
            -MemoryStartupBytes $memoryBytes `
            -SwitchName $SwitchName `
            -Path $vmRoot

        Add-VMHardDiskDrive -VMName $Name -Path $vhdPath
        Add-VMDvdDrive      -VMName $Name -Path $seedIso
        Set-VMProcessor     -VMName $Name -Count $CpuCount
        Set-VMFirmware      -VMName $Name -EnableSecureBoot Off

        $bootDisk = Get-VMHardDiskDrive -VMName $Name
        Set-VMFirmware -VMName $Name -FirstBootDevice $bootDisk

        # --- Start ------------------------------------------------------------
        Start-VM -Name $Name

        $timeoutSeconds = 30
        $elapsed        = 0
        while ((Get-VM -Name $Name).State -eq 'Off') {
            if ($elapsed -ge $timeoutSeconds) {
                Write-Warning "VM did not start within $timeoutSeconds seconds."
                break
            }
            Start-Sleep -Seconds 1
            $elapsed++
        }

        Write-Verbose "Connect with: ssh root@<vm-ip>  (password was supplied via -RootPassword)"

        Get-VM -Name $Name
    }
    finally {
        $plainPassword = $null
        if (Test-Path -LiteralPath $workDir) {
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
