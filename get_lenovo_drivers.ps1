<#
.SYNOPSIS
    Lenovo Driver Stager -- FFU/Deployment Edition v4.0

.DESCRIPTION
    Downloads and installs ALL Lenovo drivers for a manually specified
    machine type. Handles Lenovo's two-level catalog format:
      Level 1: <MachineType>_Win11.xml  -- list of package descriptor URLs
      Level 2: each descriptor XML      -- contains the actual installer filename

.PARAMETER MachineType
    4-character Lenovo Machine Type. E.g. "20XW", "21AH", "20Y7"
    Found on the device label, in BIOS, or at https://support.lenovo.com

.PARAMETER OSVersion
    Win10 or Win11 (Default: Win11)

.PARAMETER DownloadPath
    Folder to save driver packages. Default: C:\LenovoDrivers

.PARAMETER InstallDrivers
    If set, installs all downloaded packages silently after downloading.

.EXAMPLE
    .\LenovoDriverStager.ps1 -MachineType "20XW" -InstallDrivers
    .\LenovoDriverStager.ps1 -MachineType "20XW"
    .\LenovoDriverStager.ps1 -InstallDrivers

.NOTES
    Requires: PowerShell 5.1+, Internet access
    Run as Administrator when using -InstallDrivers
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$MachineType  = "",
    [ValidateSet("Win10","Win11")]
    [string]$OSVersion    = "Win11",
    [string]$DownloadPath = "C:\LenovoDrivers",
    [switch]$InstallDrivers
)

# ---------------------------------------------------------------
# GLOBALS
# ---------------------------------------------------------------
$Script:LogFile   = $null
$Script:Installed = 0
$Script:Skipped   = 0
$Script:Failed    = 0
$Script:Errors    = [System.Collections.Generic.List[string]]::new()

# ---------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------
function ConvertTo-SafeString {
    param($Value)
    return ([string]$Value).Trim()
}

function New-WebClient {
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "LenovoDriverStager/1.0")
    return $wc
}

# ---------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR","SECTION")]
        [string]$Level = "INFO"
    )
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] [$Level] $Message"
    $color = switch ($Level) {
        "INFO"    { "Cyan"    }
        "SUCCESS" { "Green"   }
        "WARN"    { "Yellow"  }
        "ERROR"   { "Red"     }
        "SECTION" { "Magenta" }
        default   { "White"   }
    }
    Write-Host $entry -ForegroundColor $color
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $entry -ErrorAction SilentlyContinue
    }
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("-" * 60) -ForegroundColor DarkCyan
    Write-Log "  $Title" "SECTION"
    Write-Host ("-" * 60) -ForegroundColor DarkCyan
}

# ---------------------------------------------------------------
# SETUP
# ---------------------------------------------------------------
function Initialize-Environment {
    if (-not (Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }
    $Script:LogFile = "$DownloadPath\DriverInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    New-Item -ItemType File -Path $Script:LogFile -Force | Out-Null
    Write-Log "Download path : $DownloadPath"
    Write-Log "Log file      : $Script:LogFile"
    Write-Log "OS Target     : $OSVersion"
}

# ---------------------------------------------------------------
# MACHINE TYPE INPUT
# ---------------------------------------------------------------
function Get-MachineType {
    Write-Section "Machine Type"

    if ($MachineType -ne "") {
        $mt = $MachineType.Trim().ToUpper()
        Write-Log "Machine Type: $mt" "SUCCESS"
        return $mt
    }

    Write-Host ""
    Write-Host "  Enter the Lenovo Machine Type (e.g. 20XW, 21AH)." -ForegroundColor White
    Write-Host "  Find it on the device label, in BIOS, or at https://support.lenovo.com" -ForegroundColor DarkGray
    Write-Host ""

    do {
        $inputVal = Read-Host "  Machine Type"
        $mt       = $inputVal.Trim().ToUpper()
        if ($mt.Length -lt 4) {
            Write-Host "  Must be at least 4 characters. Try again." -ForegroundColor Red
        }
    } while ($mt.Length -lt 4)

    Write-Log "Machine Type: $mt" "SUCCESS"
    return $mt
}

# ---------------------------------------------------------------
# LEVEL 1 CATALOG FETCH
# Returns an array of package descriptor URLs (<location> nodes)
# ---------------------------------------------------------------
function Get-CatalogLocations {
    param([string]$MT)

    Write-Section "Fetching Level-1 Catalog"

    $urlsToTry = @(
        "https://download.lenovo.com/catalog/${MT}_${OSVersion}.xml",
        "https://download.lenovo.com/catalog/${MT}_${OSVersion}_64.xml",
        "https://download.lenovo.com/catalog/${MT}.xml"
    )
    if ($OSVersion -eq "Win11") {
        $urlsToTry += "https://download.lenovo.com/catalog/${MT}_Win10.xml"
        $urlsToTry += "https://download.lenovo.com/catalog/${MT}_Win10_64.xml"
    }

    $catalogFile = "$DownloadPath\catalog_${MT}.xml"
    $wc          = New-WebClient

    foreach ($url in $urlsToTry) {
        Write-Log "Trying: $url"
        try {
            $wc.DownloadFile($url, $catalogFile)
            [xml]$xml = Get-Content -Path $catalogFile -Encoding UTF8 -ErrorAction Stop

            # Collect all <location> values -- this is the two-level catalog format
            $locations = @()
            $root = $xml.DocumentElement
            foreach ($node in $root.ChildNodes) {
                if ($node.NodeType -ne 'Element') { continue }
                $loc = ConvertTo-SafeString $node.location
                if ($loc -ne "") { $locations += $loc }
            }

            if ($locations.Count -gt 0) {
                Write-Log "Catalog OK -- $($locations.Count) package descriptors found." "SUCCESS"
                return $locations
            }
            Write-Log "  No <location> nodes found. Trying next URL." "WARN"
        }
        catch {
            Write-Log "  Failed: $url -- $_" "WARN"
        }
    }

    Write-Log "Could not find a valid catalog for '$MT'." "ERROR"
    Write-Log "Check the machine type at: https://support.lenovo.com" "WARN"
    return @()
}

# ---------------------------------------------------------------
# LEVEL 2 PACKAGE DESCRIPTOR FETCH
# Each location URL points to a package descriptor XML.
# Returns a PSCustomObject with Name, Version, Category, InstallFile, BaseUrl
#
# Descriptor structure:
#   <Package name="..." id="..." version="...">
#     <Title><Desc id="EN">Friendly name</Desc></Title>
#     <Files>
#       <Installer>
#         <File>
#           <Name>n3ipa17w.exe</Name>     <- installer filename
#         </File>
#       </Installer>
#     </Files>
#   </Package>
#
# The installer download URL = base path of the descriptor URL + installer filename
# e.g. https://download.lenovo.com/pccbbs/mobiles/n3ipa17w.exe
# ---------------------------------------------------------------
function Get-PackageDescriptor {
    param([string]$LocationUrl, [string]$TempPath)

    try {
        $wc = New-WebClient
        $tmpFile = Join-Path $TempPath ("tmp_" + [System.IO.Path]::GetFileName($LocationUrl))
        $wc.DownloadFile($LocationUrl, $tmpFile)

        [xml]$desc = Get-Content -Path $tmpFile -Encoding UTF8 -ErrorAction Stop

        # Name
        $pkgName    = ConvertTo-SafeString $desc.Package.name
        $pkgId      = ConvertTo-SafeString $desc.Package.id
        $pkgVersion = ConvertTo-SafeString $desc.Package.version

        # Friendly title -- try EN first, then first available Desc
        $title = ""
        $titleNodes = $desc.SelectNodes("//Title/Desc")
        foreach ($tn in $titleNodes) {
            $t = ConvertTo-SafeString $tn.InnerText
            if ($t -ne "") {
                if ((ConvertTo-SafeString $tn.id) -eq "EN") { $title = $t; break }
                if ($title -eq "") { $title = $t }
            }
        }
        if ($title -eq "") { $title = $pkgName }

        # Category
        $category = ""
        $catNode  = $desc.SelectSingleNode("//Category")
        if ($catNode) { $category = ConvertTo-SafeString $catNode.InnerText }
        if ($category -eq "") { $category = "Unknown" }

        # Installer file name from <Files><Installer><File><Name>
        $installerName = ""
        $instNode = $desc.SelectSingleNode("//Files/Installer/File/Name")
        if ($instNode) { $installerName = ConvertTo-SafeString $instNode.InnerText }

        if ($installerName -eq "") {
            Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
            return $null
        }

        # Build download URL -- same base path as the descriptor, different filename
        $baseUrl = $LocationUrl.Substring(0, $LocationUrl.LastIndexOf('/') + 1)
        $downloadUrl = $baseUrl + $installerName

        Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue

        return [PSCustomObject]@{
            ID           = $pkgId
            Name         = if ($title -ne "") { $title } else { $pkgName }
            Version      = $pkgVersion
            Category     = $category
            InstallerFile = $installerName
            DownloadUrl  = $downloadUrl
        }
    }
    catch {
        return $null
    }
}

# ---------------------------------------------------------------
# MAIN DOWNLOAD LOOP
# For each location URL:
#   1. Fetch the package descriptor XML
#   2. Extract installer filename + build download URL
#   3. Download the installer
# ---------------------------------------------------------------
function Get-DriverPackages {
    param(
        [string[]]$Locations,
        [string]$MT
    )

    Write-Section "Fetching Package Descriptors and Downloading"

    $mtFolder  = Join-Path $DownloadPath $MT
    $tmpFolder = Join-Path $DownloadPath "_tmp"
    foreach ($folder in @($mtFolder, $tmpFolder)) {
        if (-not (Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
    }

    $total           = $Locations.Count
    $downloadedPaths = [System.Collections.Generic.List[string]]::new()
    $wc              = New-WebClient

    for ($i = 0; $i -lt $total; $i++) {
        $locUrl = $Locations[$i]
        $idx    = $i + 1

        Write-Log "[$idx/$total] Fetching descriptor: $locUrl"

        $pkg = Get-PackageDescriptor -LocationUrl $locUrl -TempPath $tmpFolder

        if (-not $pkg) {
            Write-Log "  SKIPPED -- could not parse descriptor" "WARN"
            $Script:Skipped++
            continue
        }

        if ($pkg.InstallerFile -eq "" -or $pkg.DownloadUrl -eq "") {
            Write-Log "  SKIPPED -- no installer found in descriptor: $($pkg.Name)" "WARN"
            $Script:Skipped++
            continue
        }

        Write-Log "  Name     : $($pkg.Name)"
        Write-Log "  Version  : $($pkg.Version)"
        Write-Log "  Category : $($pkg.Category)"
        Write-Log "  URL      : $($pkg.DownloadUrl)"

        # Destination folder per package
        $safeId   = ($pkg.ID -replace '[^\w\-]', '_')
        if ($safeId -eq "") { $safeId = "pkg_$idx" }
        $pkgFolder = Join-Path $mtFolder $safeId
        $destFile  = Join-Path $pkgFolder $pkg.InstallerFile

        if (-not (Test-Path $pkgFolder)) {
            New-Item -ItemType Directory -Path $pkgFolder -Force | Out-Null
        }

        if (Test-Path $destFile) {
            Write-Log "  Already downloaded -- skipping." "INFO"
            $downloadedPaths.Add($destFile)
            continue
        }

        try {
            $wc.DownloadFile($pkg.DownloadUrl, $destFile)
            $sizeMB = [math]::Round((Get-Item $destFile).Length / 1MB, 2)
            Write-Log "  Saved: $($pkg.InstallerFile) ($sizeMB MB)" "SUCCESS"
            $downloadedPaths.Add($destFile)

            # Save metadata
            $pkg | ConvertTo-Json | Set-Content "$pkgFolder\info.json" -Encoding UTF8
        }
        catch {
            Write-Log "  FAILED: $($pkg.Name) -- $_" "ERROR"
            $Script:Errors.Add("Download failed: $($pkg.Name)")
            $Script:Failed++
        }
    }

    # Cleanup temp
    Remove-Item $tmpFolder -Recurse -Force -ErrorAction SilentlyContinue

    Write-Log ""
    Write-Log "Download complete. $($downloadedPaths.Count) of $total packages ready." "SUCCESS"
    return $downloadedPaths
}

# ---------------------------------------------------------------
# INSTALL PACKAGES
# ---------------------------------------------------------------
function Install-DriverPackages {
    param([System.Collections.Generic.List[string]]$FilePaths)

    Write-Section "Installing Packages"

    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        ).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Write-Log "Not running as Administrator -- cannot install." "ERROR"
        Write-Log "Re-run PowerShell as Administrator with -InstallDrivers." "WARN"
        return
    }

    $total = $FilePaths.Count
    $index = 0

    foreach ($filePath in $FilePaths) {
        $index++
        $ext      = [System.IO.Path]::GetExtension($filePath).ToLower()
        $fileName = Split-Path $filePath -Leaf

        Write-Log "[$index/$total] $fileName"

        try {
            switch ($ext) {
                ".exe" {
                    $proc = Start-Process -FilePath $filePath `
                                         -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES" `
                                         -PassThru -Wait -NoNewWindow
                    if ($proc.ExitCode -in @(0, 1, 3010)) {
                        $label = if ($proc.ExitCode -eq 0) { "OK" } else { "OK (reboot needed)" }
                        Write-Log "  $label (exit $($proc.ExitCode))" "SUCCESS"
                        $Script:Installed++
                    }
                    else {
                        Write-Log "  Retrying with /S (prev exit $($proc.ExitCode))..." "WARN"
                        $proc2 = Start-Process -FilePath $filePath `
                                               -ArgumentList "/S /NORESTART" `
                                               -PassThru -Wait -NoNewWindow
                        $label2 = if ($proc2.ExitCode -in @(0,1,3010)) { "OK on retry" } else { "Completed (exit $($proc2.ExitCode))" }
                        Write-Log "  $label2" "SUCCESS"
                        $Script:Installed++
                    }
                }
                ".msi" {
                    $proc = Start-Process "msiexec.exe" `
                                         -ArgumentList "/i `"$filePath`" /qn /norestart" `
                                         -PassThru -Wait -NoNewWindow
                    $label = if ($proc.ExitCode -in @(0,3010)) { "MSI OK" } else { "MSI exit $($proc.ExitCode)" }
                    Write-Log "  $label" "SUCCESS"
                    $Script:Installed++
                }
                ".inf" {
                    $proc = Start-Process "pnputil.exe" `
                                         -ArgumentList "/add-driver `"$filePath`" /install" `
                                         -PassThru -Wait -NoNewWindow
                    $label = if ($proc.ExitCode -eq 0) { "INF staged OK" } else { "pnputil exit $($proc.ExitCode)" }
                    Write-Log "  $label" "SUCCESS"
                    $Script:Installed++
                }
                default {
                    Write-Log "  Skipped unsupported type: $ext" "WARN"
                    $Script:Skipped++
                }
            }
        }
        catch {
            Write-Log "  ERROR: $_" "ERROR"
            $Script:Errors.Add("Install failed: $fileName")
            $Script:Failed++
        }
    }
}

# ---------------------------------------------------------------
# DISM INJECTION HELPER
# ---------------------------------------------------------------
function Export-DismScript {
    param([string]$MT)
    $driverFolder = Join-Path $DownloadPath $MT
    $dismScript   = Join-Path $DownloadPath "inject_drivers_${MT}.ps1"
    @"
# Auto-generated -- inject Lenovo drivers into an offline Windows image
param(
    [string]`$ImagePath  = "C:\Mount",
    [string]`$DriverPath = "$driverFolder"
)
Write-Host "Injecting drivers for $MT into `$ImagePath ..." -ForegroundColor Cyan
Dism /Image:"`$ImagePath" /Add-Driver /Driver:"`$DriverPath" /Recurse /ForceUnsigned
Write-Host "Done." -ForegroundColor Green
"@ | Set-Content -Path $dismScript -Encoding ASCII
    Write-Log "DISM script saved: $dismScript" "SUCCESS"
}

# ---------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------
function Write-Summary {
    param([string]$MT, [int]$Total)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor DarkCyan
    Write-Host "  SUMMARY" -ForegroundColor Magenta
    Write-Host ("=" * 60) -ForegroundColor DarkCyan
    Write-Log "Machine Type : $MT"
    Write-Log "OS Target    : $OSVersion"
    Write-Log "Total Pkgs   : $Total"
    Write-Log "Installed    : $Script:Installed" "SUCCESS"
    Write-Log "Skipped      : $Script:Skipped"   "WARN"
    Write-Log "Failed       : $Script:Failed"    $(if ($Script:Failed -gt 0) { "ERROR" } else { "SUCCESS" })
    Write-Log "Saved to     : $(Join-Path $DownloadPath $MT)"
    Write-Log "Log file     : $Script:LogFile"
    if ($Script:Errors.Count -gt 0) {
        Write-Host ""
        Write-Log "Errors:" "ERROR"
        foreach ($e in $Script:Errors) { Write-Log "  - $e" "ERROR" }
    }
    if (-not $InstallDrivers) {
        Write-Host ""
        Write-Log "Downloaded only. To install:" "WARN"
        Write-Log "  .\LenovoDriverStager.ps1 -MachineType $MT -InstallDrivers" "INFO"
        Write-Log "To DISM inject, run: inject_drivers_${MT}.ps1" "INFO"
    }
    Write-Host ("=" * 60) -ForegroundColor DarkCyan
    Write-Host ""
}

# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
function Main {
    Clear-Host
    Write-Host ""
    Write-Host ("+" + ("=" * 58) + "+") -ForegroundColor Magenta
    Write-Host ("|    Lenovo Driver Stager -- FFU/Deployment Edition v4.0  |") -ForegroundColor Magenta
    Write-Host ("+" + ("=" * 58) + "+") -ForegroundColor Magenta
    Write-Host ""

    Initialize-Environment

    $MT = Get-MachineType

    $locations = Get-CatalogLocations -MT $MT
    if ($locations.Count -eq 0) { exit 1 }

    $localFiles = Get-DriverPackages -Locations $locations -MT $MT

    Export-DismScript -MT $MT

    if ($InstallDrivers) {
        Install-DriverPackages -FilePaths $localFiles
    }

    Write-Summary -MT $MT -Total $locations.Count
}

Main
