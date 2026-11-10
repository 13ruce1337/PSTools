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
          - Windows ADK "Deployment Tools" (oscdimg.exe); auto-installed via
            winget (Microsoft.WindowsADK) if missing
          - winget (needed if oscdimg.exe or QEMU is not already installed)

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
        Path to oscdimg.exe. If it does not exist, the Windows ADK is
        installed automatically via winget before this path is re-checked.

    .PARAMETER IsoBootDelaySeconds
        Seconds to wait after the first boot reaches Running, before the
        one-time reboot that lets cloud-init consume the NoCloud seed. Default 4.
        Set to 0 to reboot as soon as the VM is Running.

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
        [string]$OscdimgPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe',

        [ValidateRange(0, 600)]
        [int]$IsoBootDelaySeconds = 4
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

    # Returns the full path to oscdimg.exe, installing the Windows ADK via
    # winget if it is not already available. Mirrors Get-QemuImgPath.
    function Get-OscdimgPath {
        param([string]$PreferredPath)

        # 1. Use the caller-specified / default ADK path if it is present.
        if ($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) {
            return $PreferredPath
        }

        # 2. Look for oscdimg.exe anywhere already on PATH.
        $found = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
        if ($found) {
            return $found.Source
        }

        # 3. Not found anywhere: install the Windows ADK (Deployment Tools) via winget.
        if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
            throw 'oscdimg.exe (Windows ADK Deployment Tools) is not installed and winget is ' +
                  'not available to install it. Install it manually from: ' +
                  'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install'
        }

        Write-Verbose 'oscdimg.exe not found. Installing the Windows ADK (Deployment Tools) via winget...'
        & winget.exe install `
            --id Microsoft.WindowsADK `
            --accept-source-agreements `
            --accept-package-agreements `
            --silent | Out-Host

        if ($LASTEXITCODE -ne 0) {
            Write-Warning "winget exited with code $LASTEXITCODE; checking whether oscdimg.exe is available anyway."
        }

        # Refresh PATH from machine and user environment.
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                    [System.Environment]::GetEnvironmentVariable('Path', 'User')

        # 4. Re-check the preferred path and PATH after install.
        if ($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) {
            return $PreferredPath
        }
        $found = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
        if ($found) {
            return $found.Source
        }

        # 5. Fall back to the default ADK install location (search all versions).
        $adkRoot = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg'
        if (Test-Path -LiteralPath $adkRoot) {
            $candidate = Get-ChildItem -LiteralPath $adkRoot -Recurse -Filter 'oscdimg.exe' -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($candidate) {
                return $candidate.FullName
            }
        }

        throw 'The Windows ADK was installed but oscdimg.exe could not be located. ' +
              'Install the "Deployment Tools" feature manually: ' +
              'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install'
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

    # Resolve (and, if needed, auto-install) oscdimg.exe. Get-OscdimgPath
    # installs the Windows ADK via winget when the tool is missing, then
    # verifies the result. It throws only if the tool is genuinely unavailable.
    try {
        $OscdimgExe = Get-OscdimgPath -PreferredPath $OscdimgPath
    }
    catch {
        throw ("oscdimg.exe could not be located or auto-installed. " +
               "Last error: $($_.Exception.Message)")
    }
    Write-Verbose "Using oscdimg.exe at: $OscdimgExe"

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

        # Use ISO9660 (-l) with UDF (-udf). cloud-init's NoCloud source reads the
        # seed via the CD-ROM device (/dev/sr0), and ISO9660+UDF is the most
        # reliably handled by its reader on a fresh boot. (Avoid -j1/Joliet-only:
        # some Linux CD-ROM stacks mis-handle long filenames there, which can
        # cause the seed to be missed on first boot -- the classic "root still
        # locked, boot stalls, reboot fixes it" symptom.)
        $oscdimgArgs = @('-udf', '-l', '-m', 'RECYCLER', '-lcidata', $workDir, $seedIso)
        & $OscdimgExe @oscdimgArgs | Out-Null
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
        # A freshly created VM boots "fine" from Hyper-V's point of view, but the
        # Debian cloud image's cloud-init may miss the NoCloud seed on its very
        # first boot (boot-order / CD-ROM readiness race). The guest then stalls
        # part-way through boot with root still locked -- and a cold reboot fixes
        # it, because cloud-init picks up the seed on the second pass. We can't
        # see the guest error from the host (it's inside the VM), so we reproduce
        # the reboot that reliably works.
        function Wait-VmState {
            param(
                [Parameter(Mandatory)][ValidateSet('Running', 'Off')][string]$TargetState,
                [int]$Seconds = 60
            )
            $deadline = (Get-Date).AddSeconds($Seconds)
            do {
                Start-Sleep -Seconds 1
                $vm = Get-VM -Name $Name -ErrorAction SilentlyContinue
                if ($vm -and $vm.State -eq $TargetState) { return $true }
            } while ((Get-Date) -lt $deadline)
            return $false
        }

        Write-Verbose 'Starting VM (first boot)...'
        Start-VM -Name $Name

        if (-not (Wait-VmState -TargetState 'Running' -Seconds 60)) {
            $vm     = Get-VM -Name $Name -ErrorAction SilentlyContinue
            $state  = $vm.State
            $reason = $vm.Error
            throw "VM '$Name' did not reach a running state (state: $state). $reason"
        }

        # Give the guest a moment to begin first-boot cloud-init, then cold-reboot
        # so cloud-init reliably consumes the NoCloud seed (the reboot you were
        # doing by hand). This is the reliable fix for the "boots but fails to
        # fully start on first boot" symptom.
        if ($IsoBootDelaySeconds -gt 0) {
            Write-Verbose "First boot reached Running; waiting ${IsoBootDelaySeconds}s before a one-time reboot to let cloud-init finish, then re-booting to guarantee the seed is applied."
            Start-Sleep -Seconds $IsoBootDelaySeconds
        }
        else {
            Write-Verbose 'First boot reached Running; performing a one-time reboot to guarantee cloud-init consumes the NoCloud seed.'
        }

        Write-Verbose 'Rebooting VM once so cloud-init applies the NoCloud seed (root password, hostname, networking)...'
        try {
            Reset-VM -Name $Name -Force
        }
        catch {
            # If the guest is mid-boot and won't accept a graceful reset, a hard
            # restart is equivalent for our purposes.
            Write-Verbose 'Reset-VM did not complete; using a hard Restart-VM instead.'
            Stop-VM  -Name $Name -Force -ErrorAction SilentlyContinue
            Start-VM -Name $Name
        }

        if (-not (Wait-VmState -TargetState 'Running' -Seconds 60)) {
            $vm     = Get-VM -Name $Name -ErrorAction SilentlyContinue
            $state  = $vm.State
            $reason = $vm.Error
            throw "VM '$Name' did not come up after the reboot (state: $state). $reason " +
                  'Open the VM console in Hyper-V Manager and check the guest boot output.'
        }

        Write-Verbose "VM '$Name' is running and cloud-init has been given its reboot."
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
