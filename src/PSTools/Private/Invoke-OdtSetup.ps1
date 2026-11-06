function Invoke-OdtSetup {
    <#
    .SYNOPSIS
        Runs 'setup.exe /configure' with the supplied ODT configuration XML.

    .DESCRIPTION
        Writes the configuration to a uniquely named temp file next to
        setup.exe, runs it, removes the file and returns the exit code.
        Throws when setup.exe exits non-zero.

        Waits only for setup.exe itself. Start-Process -Wait would also wait
        for every process setup.exe starts, and some Click-to-Run processes
        keep running after the install finishes, so -Wait never returns.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Configuration,

        # Folder ODT writes logs to (used for directory creation and error text).
        [string]$LogPath,

        # Path to setup.exe. Resolved/downloaded automatically if omitted.
        [string]$SetupPath,

        # Close running Office applications instead of failing.
        [switch]$Force
    )

    if (-not $SetupPath) {
        $SetupPath = Get-OdtSetupPath
    }

    Close-OfficeApplication -Force:$Force

    if ($LogPath -and -not (Test-Path -LiteralPath $LogPath)) {
        $null = New-Item -ItemType Directory -Path $LogPath -Force
    }

    $configPath = Join-Path (Split-Path -Path $SetupPath -Parent) ('config_{0}.xml' -f [guid]::NewGuid().ToString('N'))
    $utf8NoBom  = New-Object System.Text.UTF8Encoding($false)

    try {
        [System.IO.File]::WriteAllText($configPath, $Configuration, $utf8NoBom)
        Write-Verbose "ODT configuration:`n$Configuration"
        Write-Verbose "Running $SetupPath /configure $configPath (this may take a while)..."

        $proc = Start-Process -FilePath $SetupPath `
            -ArgumentList '/configure', "`"$configPath`"" `
            -PassThru -WindowStyle Hidden

        $null = $proc.Handle      # keeps ExitCode available in Windows PowerShell 5.1
        $proc.WaitForExit()       # waits for setup.exe only, not its child processes

        if ($proc.ExitCode -ne 0) {
            $hint = if ($LogPath) { " See the ODT logs in '$LogPath'." } else { '' }
            throw "ODT setup.exe exited with code $($proc.ExitCode).$hint"
        }

        $proc.ExitCode
    }
    finally {
        if (Test-Path -LiteralPath $configPath) {
            Remove-Item -LiteralPath $configPath -Force -ErrorAction SilentlyContinue
        }
    }
}
