# Last-known-good ODT bootstrapper. Used only if the current link can't be
# resolved from Microsoft's download page.
$script:OdtFallbackUrl = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17126-20132.exe'
$script:OdtDownloadPage = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'

function Get-OdtSetupPath {
    <#
    .SYNOPSIS
        Returns the path to the Office Deployment Tool setup.exe, downloading
        and extracting it if needed.

    .DESCRIPTION
        setup.exe is cached in -WorkPath and reused on later calls. Pass
        -Refresh to force a fresh download.

        If -OdtUrl is not supplied, the current bootstrapper link is scraped
        from Microsoft's download page, falling back to a pinned URL. The
        bootstrapper's Authenticode signature must be valid and issued to
        Microsoft before it is executed.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string]$OdtUrl,

        [ValidateNotNullOrEmpty()]
        [string]$WorkPath = (Join-Path $env:TEMP 'PSTools\ODT'),

        [switch]$Refresh
    )

    $setupExe = Join-Path $WorkPath 'setup.exe'

    if (-not $Refresh -and (Test-Path -LiteralPath $setupExe)) {
        Write-Verbose "Using cached ODT: $setupExe"
        return $setupExe
    }

    if (-not (Test-Path -LiteralPath $WorkPath)) {
        $null = New-Item -ItemType Directory -Path $WorkPath -Force
    }

    if (Test-Path -LiteralPath $setupExe) {
        Remove-Item -LiteralPath $setupExe -Force
    }

    # --- Resolve download URL -----------------------------------------------
    if (-not $OdtUrl) {
        try {
            Write-Verbose "Resolving current ODT download link from $script:OdtDownloadPage"
            $page = Invoke-WebRequest -Uri $script:OdtDownloadPage -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            $match = [regex]::Match(
                $page.Content,
                'https://download\.microsoft\.com/download/[^"''\s<>]+?/officedeploymenttool_[\d-]+\.exe'
            )
            if ($match.Success) {
                $OdtUrl = $match.Value
            }
        }
        catch {
            Write-Verbose "Could not resolve current ODT link: $($_.Exception.Message)"
        }

        if (-not $OdtUrl) {
            Write-Verbose 'Falling back to pinned ODT URL.'
            $OdtUrl = $script:OdtFallbackUrl
        }
    }

    $bootstrapper = Join-Path $WorkPath ([System.IO.Path]::GetFileName(([uri]$OdtUrl).LocalPath))

    try {
        # --- Download -------------------------------------------------------
        Write-Verbose "Downloading ODT from $OdtUrl"
        $oldProgress = $ProgressPreference
        try {
            # The progress bar makes Invoke-WebRequest extremely slow in Windows PowerShell 5.1.
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $OdtUrl -OutFile $bootstrapper -UseBasicParsing -ErrorAction Stop
        }
        finally {
            $ProgressPreference = $oldProgress
        }

        # --- Verify signature -----------------------------------------------
        $sig = Get-AuthenticodeSignature -FilePath $bootstrapper
        if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
            throw "ODT bootstrapper signature check failed (status: $($sig.Status), signer: $($sig.SignerCertificate.Subject))."
        }

        # --- Extract --------------------------------------------------------
        Write-Verbose "Extracting ODT to $WorkPath"
        $proc = Start-Process -FilePath $bootstrapper `
            -ArgumentList '/quiet', "/extract:`"$WorkPath`"" `
            -Wait -PassThru -WindowStyle Hidden

        if ($proc.ExitCode -ne 0) {
            throw "ODT extraction failed with exit code $($proc.ExitCode)."
        }

        if (-not (Test-Path -LiteralPath $setupExe)) {
            throw "ODT extraction completed but setup.exe was not found in $WorkPath."
        }
    }
    finally {
        if (Test-Path -LiteralPath $bootstrapper) {
            Remove-Item -LiteralPath $bootstrapper -Force -ErrorAction SilentlyContinue
        }
    }

    $setupExe
}
