function New-DebianVM {
    <#
    .SYNOPSIS
        Creates and starts a Generation 2 Hyper-V VM from a Debian cloud image.

    .DESCRIPTION
        Copies a base Debian cloud VHDX, resizes it, builds a cloud-init NoCloud
        seed ISO with oscdimg (Windows ADK Deployment Tools), creates the VM,
        attaches both, and starts it.

        Requirements:
          - Elevated session with the Hyper-V PowerShell module
          - Windows ADK "Deployment Tools" (oscdimg.exe)
          - A Debian cloud image converted to VHDX (see the Debian VHDX helper)

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
        Path to the base Debian VHDX.

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

    .OUTPUTS
        Microsoft.HyperV.PowerShell.VirtualMachine
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([Microsoft.HyperV.PowerShell.VirtualMachine])]
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
        [string]$SeedsPath = (Join-Path $env:USERPROFILE 'Seeds'),

        [ValidateNotNullOrEmpty()]
        [string]$SwitchName = 'Default Switch',

        [ValidateNotNullOrEmpty()]
        [string]$OscdimgPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe'
    )

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

    if (-not (Test-Path -LiteralPath $BaseVhdxPath)) {
        throw "Base image not found at '$BaseVhdxPath'. Create it first (see create_debian_vhdx) and place it there."
    }

    if (-not (Test-Path -LiteralPath $OscdimgPath)) {
        throw ("oscdimg.exe not found at '$OscdimgPath'. Install the Windows ADK 'Deployment Tools' feature: " +
               'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install')
    }

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create and start Debian VM')) {
        return
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
