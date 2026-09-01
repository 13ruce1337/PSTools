#requires -RunAsAdministrator
<#
.SYNOPSIS
    Stages Lenovo driver packs (SCCM + discrete GFX) into the local driver
    store of a Hyper-V staging VM, prior to sysprep / FFU capture.

.DESCRIPTION
    Reads Lenovo's live catalog (catalogv2.xml), resolves the correct pack for a
    machine type (MTM), downloads with hash verification and caching, silently
    extracts, stages every INF with pnputil, then writes a JSON manifest into
    C:\Windows\Logs\ImageDrivers so the resulting FFU is self-documenting.

    v2 changes:
      - BOM-safe catalog parsing via XmlDocument.Load() (fixes the "Data at the
        root level is invalid" dump)
      - Handles <GFX> discrete graphics packs (NVIDIA/AMD on P-series)
      - Ignores <HSA> Hardware Support Apps (Store apps, not INF drivers)
      - Dedupes packs that share a single URL across multiple OS versions
      - All errors truncated so a parse failure never dumps 1.4 MB to console

.EXAMPLE
    .\Add-LenovoDriversToImage.ps1 -MachineType 21HD -OSVersion 25H2

.EXAMPLE
    .\Add-LenovoDriversToImage.ps1 -ModelName 'X1 Carbon Gen 12' -Prune

.EXAMPLE
    '21HD','21K3','21F6' | ForEach-Object {
        .\Add-LenovoDriversToImage.ps1 -MachineType $_ -OSVersion 25H2
    }
#>
[CmdletBinding(DefaultParameterSetName = 'ByType')]
param(
    # 4-character Lenovo machine type, e.g. 21HD. This is the reliable key.
    [Parameter(Mandatory, ParameterSetName = 'ByType', Position = 0)]
    [ValidatePattern('^[A-Za-z0-9]{4}$')]
    [string]$MachineType,

    # Substring of the catalog model name. Use only if you don't know the MTM.
    [Parameter(Mandatory, ParameterSetName = 'ByName', Position = 0)]
    [string]$ModelName,

    [ValidateSet('21H2','22H2','23H2','24H2','25H2')]
    [string]$OSVersion,

    # Include discrete graphics packs (<GFX>). Required for P1/P16/ThinkStation.
    [switch]$IncludeGraphics,

    # Drop categories that don't belong in a base image.
    [switch]$Prune,

    [string]$WorkRoot  = 'C:\_DriverStage',
    [string]$CacheRoot = 'C:\_DriverCache',

    # Keep extracted INF trees on disk (debugging).
    [switch]$KeepFiles,

    # Resolve and report only. No download, no staging.
    [switch]$WhatIfOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference   = 'Stop'
$ProgressPreference      = 'SilentlyContinue'   # ~10x faster web transfers
$FormatEnumerationLimit  = 4                    # never carpet-bomb the console

$CatalogUrl = 'https://download.lenovo.com/cdrt/td/catalogv2.xml'
$LogDir     = 'C:\Windows\Logs\ImageDrivers'
$FallbackOrder = @('25H2','24H2','23H2','22H2','21H2')

$null = New-Item $LogDir -ItemType Directory -Force

#region Helpers ---------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK','STEP')][string]$Level = 'INFO'
    )
    $color = @{ INFO='Gray'; WARN='Yellow'; ERROR='Red'; OK='Green'; STEP='Cyan' }[$Level]
    Write-Host ('[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message) -ForegroundColor $color
    Add-Content -LiteralPath (Join-Path $LogDir 'staging.log') -Value ('{0} [{1}] {2}' -f (Get-Date -Format s), $Level, $Message)
}

function Get-ShortError {
    # Keeps a failed [xml] cast or similar from printing the whole document.
    param([Parameter(Mandatory)]$ErrorRecord, [int]$Max = 400)
    $m = $ErrorRecord.Exception.Message -replace '\s+', ' '
    if ($m.Length -gt $Max) { $m.Substring(0, $Max) + ' ...[truncated]' } else { $m }
}

function Get-LenovoCatalog {
    param([Parameter(Mandatory)][string]$Url)

    # Lenovo serves this with a UTF-8 BOM. Invoke-WebRequest decodes the BOM to
    # the literal chars 'i>>?' which breaks a [xml] string cast. Loading from
    # bytes lets XmlDocument do its own encoding detection instead.
    $tmp = Join-Path $env:TEMP ('lenovo_catalogv2_{0}.xml' -f $PID)
    try {
        Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing -TimeoutSec 120

        $size = [Math]::Round((Get-Item $tmp).Length / 1KB, 0)
        if ($size -lt 100) { throw "Catalog download is only $size KB - looks truncated or blocked." }

        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null          # no external DTD/entity fetches
        $doc.Load($tmp)                   # byte-level, BOM-safe

        $models = @($doc.ModelList.Model)
        if ($models.Count -eq 0) { throw 'Catalog parsed but contained no <Model> nodes.' }
        return $models
    }
    catch {
        throw ('Failed to load Lenovo catalog: {0}' -f (Get-ShortError $_))
    }
    finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Resolve-LenovoModel {
    param(
        [Parameter(Mandatory)][array]$Models,
        [string]$Type,
        [string]$Name
    )
    if ($Type) {
        $t = $Type.ToUpper()
        $hits = @($Models | Where-Object {
            $_.Types -and (@($_.Types.Type) -contains $t)
        })
        $label = $t
    }
    else {
        $hits = @($Models | Where-Object { $_.name -like "*$Name*" })
        $label = $Name
    }

    if ($hits.Count -eq 0) {
        throw "No catalog entry matched '$label'. Check the MTM on the underside label or run: (Get-CimInstance Win32_ComputerSystem).Model"
    }
    if ($hits.Count -gt 1) {
        Write-Log "Ambiguous match for '$label' - $($hits.Count) candidates:" 'WARN'
        foreach ($h in $hits) {
            Write-Log ('    {0}   [{1}]  arch={2}' -f $h.name, (@($h.Types.Type) -join ','), $h.arch) 'WARN'
        }
        throw "Ambiguous. Re-run with a specific -MachineType (Intel and AMD variants of the same model name are separate packs)."
    }
    return $hits[0]
}

function Select-DriverPack {
    <#  Picks one pack node for the requested OS version, walking backwards
        through the fallback order if the exact version isn't published.
        Lenovo often points several <SCCM> version rows at one file, so callers
        must dedupe on URL. #>
    param(
        [Parameter(Mandatory)]$ModelNode,
        [Parameter(Mandatory)][string]$ElementName,   # 'SCCM' or 'GFX'
        [Parameter(Mandatory)][string]$Wanted
    )
    $all = @($ModelNode.SelectNodes($ElementName) | Where-Object { $_.os -eq 'win11' })
    if ($all.Count -eq 0) { return $null }

    $exact = @($all | Where-Object { $_.version -eq $Wanted })
    if ($exact.Count -gt 0) { return $exact[0] }

    $start = [Array]::IndexOf($FallbackOrder, $Wanted)
    if ($start -lt 0) { $start = 0 }
    for ($i = $start; $i -lt $FallbackOrder.Count; $i++) {
        $c = @($all | Where-Object { $_.version -eq $FallbackOrder[$i] })
        if ($c.Count -gt 0) {
            Write-Log "No win11 $Wanted $ElementName pack; falling back to $($FallbackOrder[$i])." 'WARN'
            return $c[0]
        }
    }
    return $all[0]
}

function Get-VerifiedPack {
    <#  Download with cache reuse + SHA256 verification. The catalog's 'crc'
        attribute is actually a SHA256 despite the name. #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [string]$Sha256,
        [Parameter(Mandatory)][string]$Destination
    )
    $null = New-Item (Split-Path $Destination) -ItemType Directory -Force

    if (Test-Path $Destination) {
        if ($Sha256) {
            if ((Get-FileHash $Destination -Algorithm SHA256).Hash -eq $Sha256.ToUpper()) {
                Write-Log 'Cached copy verified - skipping download.' 'OK'
                return $Destination
            }
            Write-Log 'Cached copy failed hash check - re-downloading.' 'WARN'
        }
        Remove-Item $Destination -Force
    }

    Write-Log 'Downloading (packs are typically 1-3 GB)...'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        Start-BitsTransfer -Source $Url -Destination $Destination -Description 'Lenovo driver pack' -ErrorAction Stop
    }
    catch {
        Write-Log "BITS unavailable ($(Get-ShortError $_ 120)) - using Invoke-WebRequest." 'WARN'
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -TimeoutSec 3600
    }
    $sw.Stop()

    if (-not (Test-Path $Destination)) { throw "Download produced no file: $Url" }
    $mb = [Math]::Round((Get-Item $Destination).Length / 1MB, 1)
    Write-Log ('Downloaded {0} MB in {1:n0}s ({2:n1} MB/s)' -f $mb, $sw.Elapsed.TotalSeconds, ($mb / [Math]::Max(1, $sw.Elapsed.TotalSeconds))) 'OK'

    if ($Sha256) {
        $have = (Get-FileHash $Destination -Algorithm SHA256).Hash
        if ($have -ne $Sha256.ToUpper()) {
            Remove-Item $Destination -Force
            throw "SHA256 mismatch for $(Split-Path $Url -Leaf). Expected $($Sha256.ToUpper()), got $have. File discarded."
        }
        Write-Log 'SHA256 verified.' 'OK'
    }
    else {
        Write-Log 'Catalog supplied no hash for this pack - integrity unverified.' 'WARN'
    }
    return $Destination
}

function Expand-LenovoPack {
    # Lenovo packs are Inno Setup self-extractors.
    param(
        [Parameter(Mandatory)][string]$ExePath,
        [Parameter(Mandatory)][string]$TargetDir
    )
    if (Test-Path $TargetDir) { Remove-Item $TargetDir -Recurse -Force }
    $null = New-Item $TargetDir -ItemType Directory -Force

    Write-Log "Extracting to $TargetDir ..."
    $p = Start-Process -FilePath $ExePath `
            -ArgumentList '/VERYSILENT', "/DIR=`"$TargetDir`"", '/EXTRACT="YES"' `
            -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) { Write-Log "Extractor exit code $($p.ExitCode) (often benign)." 'WARN' }

    $infs = @(Get-ChildItem $TargetDir -Filter *.inf -Recurse -File -ErrorAction SilentlyContinue)
    if ($infs.Count -eq 0) {
        throw "Extraction produced no INF files under $TargetDir. Pack may use a different extractor - try running it manually with /? to see switches."
    }
    Write-Log "Extracted $($infs.Count) INF files." 'OK'
    return $infs.Count
}

function Remove-UnwantedDrivers {
    <#  Trims categories that bloat an image without helping first boot.
        Deliberately conservative: never touches storage, network, chipset,
        graphics, audio, input or Bluetooth. #>
    param([Parameter(Mandatory)][string]$Root)

    $patterns = @('printer','fingerprint','smartcard','modem','wwan_firmware','dock_firmware','manual','doc$')
    $removed = 0
    foreach ($pat in $patterns) {
        $dirs = @(Get-ChildItem $Root -Directory -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match $pat })
        foreach ($d in $dirs) {
            if (Test-Path $d.FullName) {
                Write-Log "  pruning $($d.Name)"
                Remove-Item $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                $removed++
            }
        }
    }
    $left = @(Get-ChildItem $Root -Filter *.inf -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-Log "Pruned $removed folders; $left INFs remain."
    return $left
}

function Get-DriverStoreCount {
    @(& pnputil.exe /enum-drivers | Select-String -SimpleMatch 'Published Name').Count
}

function Add-DriversToStore {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$LogTag
    )
    $before = Get-DriverStoreCount
    Write-Log "Driver store holds $before packages. Staging (this takes several minutes)..." 'STEP'

    $out = & pnputil.exe /add-driver (Join-Path $Root '*.inf') /subdirs 2>&1
    $logPath = Join-Path $LogDir "pnputil_$LogTag.log"
    $out | Set-Content -LiteralPath $logPath -Encoding UTF8

    $after  = Get-DriverStoreCount
    $added  = $after - $before
    $failed = @($out | Select-String -Pattern 'Failed|Adding driver package failed').Count

    if ($added -le 0) {
        throw "pnputil staged 0 packages. Review $logPath"
    }
    if ($failed -gt 0) {
        Write-Log "$failed INF(s) reported failure - usually unsigned or wrong-arch. See $logPath" 'WARN'
    }
    Write-Log "Staged $added packages (store now $after)." 'OK'
    return [pscustomobject]@{ Added = $added; Total = $after; Failed = $failed; Log = $logPath }
}

#endregion --------------------------------------------------------------------

# --- Preflight ---------------------------------------------------------------

Write-Log '=== Lenovo driver staging (pre-sysprep) ===' 'STEP'

if (-not $OSVersion) {
    $OSVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion
    Write-Log "-OSVersion not supplied; using the running build: $OSVersion"
}
if ($FallbackOrder -notcontains $OSVersion) {
    Write-Log "Running build '$OSVersion' isn't a known pack version; will fall back from 25H2." 'WARN'
    $OSVersion = '25H2'
}

$freeGB = [Math]::Round((Get-PSDrive C).Free / 1GB, 1)
Write-Log "Free space on C: $freeGB GB"
if ($freeGB -lt 25) { Write-Log 'Under 25 GB free. Download + extract + driver store needs headroom.' 'WARN' }

# --- Resolve -----------------------------------------------------------------

$models = Get-LenovoCatalog -Url $CatalogUrl
Write-Log "Catalog loaded: $($models.Count) models" 'OK'

$model = Resolve-LenovoModel -Models $models -Type $MachineType -Name $ModelName
$key   = if ($MachineType) { $MachineType.ToUpper() } else { ($model.Types.Type | Select-Object -First 1) }

Write-Log "Matched: $($model.name)" 'OK'
Write-Log "  arch  : $($model.arch)"
Write-Log "  types : $(@($model.Types.Type) -join ', ')"

# Build the work list. SCCM is mandatory; GFX is opt-in. HSA is always ignored
# (those are Store/UWP Hardware Support Apps, not INF driver packages).
$targets = @()

$sccm = Select-DriverPack -ModelNode $model -ElementName 'SCCM' -Wanted $OSVersion
if (-not $sccm) { throw "No Windows 11 SCCM driver pack published for $($model.name)." }
$targets += [pscustomobject]@{ Kind = 'SCCM'; Node = $sccm; Brand = '' }

$gfxAll = @($model.SelectNodes('GFX') | Where-Object { $_.os -eq 'win11' })
if ($gfxAll.Count -gt 0) {
    if ($IncludeGraphics) {
        foreach ($brand in (@($gfxAll.brand) | Sort-Object -Unique)) {
            $g = @($gfxAll | Where-Object { $_.brand -eq $brand -and $_.version -eq $sccm.version })
            if ($g.Count -eq 0) { $g = @($gfxAll | Where-Object { $_.brand -eq $brand }) }
            $targets += [pscustomobject]@{ Kind = 'GFX'; Node = $g[0]; Brand = $brand }
        }
        Write-Log "Including $($targets.Count - 1) discrete graphics pack(s)."
    }
    else {
        Write-Log "This model publishes a discrete <GFX> pack ($((@($gfxAll.brand) | Sort-Object -Unique) -join ',')). Without -IncludeGraphics you will get Basic Display Adapter on real hardware." 'WARN'
    }
}

# Dedupe: Lenovo frequently points multiple version rows at one file.
$seen = @{}
$work = @()
foreach ($t in $targets) {
    $u = $t.Node.'#text'.Trim()
    if ($seen.ContainsKey($u)) { Write-Log "Skipping duplicate URL for $($t.Kind)."; continue }
    $seen[$u] = $true
    $work += $t
}

Write-Log "--- Plan ($($work.Count) pack(s)) ---" 'STEP'
foreach ($t in $work) {
    $u = $t.Node.'#text'.Trim()
    Write-Log ('  [{0}{1}] win11 {2}  {3}  (published {4})' -f $t.Kind, $(if ($t.Brand) { "/$($t.Brand)" }), $t.Node.version, (Split-Path $u -Leaf), $t.Node.date)
}

if ($WhatIfOnly) {
    foreach ($t in $work) { Write-Log ('URL: {0}' -f $t.Node.'#text'.Trim()) }
    Write-Log 'WhatIfOnly - nothing downloaded or staged.' 'OK'
    return
}

# --- Download / extract / stage ----------------------------------------------

$results = @()
foreach ($t in $work) {
    $url  = $t.Node.'#text'.Trim()
    $file = Split-Path $url -Leaf
    $tag  = '{0}_{1}' -f $key, $t.Kind

    Write-Log "--- Processing $($t.Kind) pack: $file ---" 'STEP'

    $exe        = Get-VerifiedPack -Url $url -Sha256 $t.Node.crc -Destination (Join-Path $CacheRoot $file)
    $extractDir = Join-Path $WorkRoot ($file -replace '\.exe$','')
    $infCount   = Expand-LenovoPack -ExePath $exe -TargetDir $extractDir

    if ($Prune -and $t.Kind -eq 'SCCM') { $infCount = Remove-UnwantedDrivers -Root $extractDir }

    $staged = Add-DriversToStore -Root $extractDir -LogTag $tag

    $results += [pscustomobject]@{
        Kind          = $t.Kind
        Brand         = $t.Brand
        PackFile      = $file
        PackVersion   = $t.Node.version
        PackDate      = $t.Node.date
        PackSha256    = $t.Node.crc
        InfFiles      = $infCount
        PackagesAdded = $staged.Added
        Failures      = $staged.Failed
    }

    if (-not $KeepFiles) {
        Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log 'Removed extraction scratch.'
    }
}

# --- Manifest ----------------------------------------------------------------

$manifest = [pscustomobject]@{
    StagedUtc     = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
    ScriptVersion = '2.0'
    Vendor        = 'Lenovo'
    Model         = $model.name
    Architecture  = $model.arch
    MachineTypes  = @($model.Types.Type)
    TargetOS      = "win11 $OSVersion"
    ImageBuild    = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').BuildLabEx
    Pruned        = [bool]$Prune
    Packs         = $results
    StoreTotal    = Get-DriverStoreCount
}
$manifestPath = Join-Path $LogDir "manifest_$key.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

if (-not $KeepFiles -and (Test-Path $WorkRoot)) {
    Remove-Item $WorkRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Log '=== COMPLETE ===' 'OK'
Write-Log ('Model: {0} | packs: {1} | driver packages added: {2}' -f $model.name, $results.Count, (($results | Measure-Object PackagesAdded -Sum).Sum)) 'OK'
Write-Log "Manifest: $manifestPath" 'OK'
Write-Log "Cache retained at $CacheRoot - delete it before capture if it lives on the image volume." 'WARN'
Write-Log 'NEXT: sysprep /generalize with PersistAllDeviceInstalls=true, or these drivers will be stripped.' 'WARN'
