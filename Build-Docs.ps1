#Requires -Version 7.2
<#
.SYNOPSIS
    Generates the PSTools command reference in docs\ from comment-based help.

.DESCRIPTION
    Imports the PSTools module from src\PSTools and uses
    Microsoft.PowerShell.PlatyPS to write one Markdown page per command to
    docs\<Category>\<Command>.md. The category is the Public\<Category>\
    subfolder the command's file lives in. Also writes docs\README.md, an
    index of all commands grouped by category.

    The docs\ folder is fully regenerated on every run. Do not hand-edit
    files in it; change the function's comment-based help instead.

    The build fails if the files in Public\ and FunctionsToExport in
    PSTools.psd1 are out of sync. It warns about commands whose help is
    incomplete (PlatyPS placeholders in the output).

    Used by .github\workflows\docs.yml, and can be run locally from any
    working directory.

.PARAMETER OutputPath
    Folder to write the docs to. Defaults to docs\ in the repository root.

.EXAMPLE
    Install-PSResource -Name Microsoft.PowerShell.PlatyPS
    ./Build-Docs.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'docs')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot     = $PSScriptRoot
$moduleRoot   = Join-Path $repoRoot 'src' 'PSTools'
$manifestPath = Join-Path $moduleRoot 'PSTools.psd1'
$publicPath   = Join-Path $moduleRoot 'Public'

# Friendly names for category folders. Folders not listed use their folder name.
$categoryNames = @{
    ActiveDirectory = 'Active Directory'
    FileSystem      = 'File System'
    HyperV          = 'Hyper-V'
    Intune          = 'Intune'
    M365            = 'Microsoft 365'
    Network         = 'Network'
}

# --- Dependencies -----------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerShell.PlatyPS')) {
    throw "Microsoft.PowerShell.PlatyPS is not installed. Run: Install-PSResource -Name Microsoft.PowerShell.PlatyPS"
}
Import-Module -Name 'Microsoft.PowerShell.PlatyPS' -Force

# --- Check Public files against the manifest --------------------------------

$manifest  = Import-PowerShellDataFile -Path $manifestPath
$exported  = @($manifest.FunctionsToExport)
$files     = @(Get-ChildItem -Path $publicPath -Filter '*.ps1' -File -Recurse | Sort-Object -Property FullName)
$fileNames = @($files.BaseName)

$notExported = @($fileNames | Where-Object { $_ -notin $exported })
$noFile      = @($exported  | Where-Object { $_ -notin $fileNames })

if ($notExported.Count -gt 0 -or $noFile.Count -gt 0) {
    $msg = 'Public\ and FunctionsToExport in PSTools.psd1 are out of sync.'
    if ($notExported.Count -gt 0) { $msg += " Not in FunctionsToExport: $($notExported -join ', ')." }
    if ($noFile.Count -gt 0)      { $msg += " No matching file in Public\: $($noFile -join ', ')." }
    throw $msg
}

# --- Import the module ------------------------------------------------------

Get-Module -Name 'PSTools' | Remove-Module -Force
Import-Module -Name $manifestPath -Force

# --- Generate pages ---------------------------------------------------------

$staging = Join-Path ([IO.Path]::GetTempPath()) ("PSTools-docs-" + [guid]::NewGuid().ToString('N'))
$null = New-Item -Path $staging -ItemType Directory

if (Test-Path -Path $OutputPath) {
    Remove-Item -Path $OutputPath -Recurse -Force
}
$null = New-Item -Path $OutputPath -ItemType Directory

$entries    = [System.Collections.Generic.List[object]]::new()
$incomplete = [System.Collections.Generic.List[string]]::new()

try {
    foreach ($file in $files) {
        $name = $file.BaseName

        $relative = [IO.Path]::GetRelativePath($publicPath, $file.DirectoryName)
        $category = if ($relative -eq '.') { 'General' }
                    else { $relative.Split([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)[0] }

        $command = Get-Command -Module 'PSTools' -Name $name -ErrorAction SilentlyContinue
        if (-not $command) {
            throw "Command '$name' was not found after importing PSTools. Check that '$($file.FullName)' defines a function with that name."
        }

        $null = New-MarkdownCommandHelp -CommandInfo $command -OutputFolder $staging -Force

        $generated = Get-ChildItem -Path $staging -Filter "$name.md" -File -Recurse | Select-Object -First 1
        if (-not $generated) {
            throw "PlatyPS did not produce a page for '$name'."
        }

        # Normalize line endings and drop ms.date so unchanged help produces no diff.
        $text = (Get-Content -Path $generated.FullName -Raw) -replace "`r`n", "`n"
        $text = $text -replace '(?m)^ms\.date:.*\n', ''

        if ($text -match '\{\{ Fill') {
            $incomplete.Add($name)
        }

        $categoryPath = Join-Path $OutputPath $category
        if (-not (Test-Path -Path $categoryPath)) {
            $null = New-Item -Path $categoryPath -ItemType Directory
        }
        Set-Content -Path (Join-Path $categoryPath "$name.md") -Value $text -NoNewline -Encoding utf8NoBOM

        # Synopsis for the index. Functions without help return their syntax here.
        $synopsis = ((Get-Help -Name $name).Synopsis -replace '\s+', ' ').Trim()
        if ($synopsis -match "^$([regex]::Escape($name))\b") { $synopsis = '' }
        $synopsis = $synopsis -replace '\|', '\|'

        $entries.Add([pscustomobject]@{
            Category = $category
            Name     = $name
            Synopsis = $synopsis
        })

        Write-Verbose "Generated $category/$name.md"
    }
}
finally {
    Remove-Item -Path $staging -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Index ------------------------------------------------------------------

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# PSTools command reference')
$lines.Add('')

foreach ($group in ($entries | Group-Object -Property Category | Sort-Object -Property Name)) {
    $title = if ($categoryNames.ContainsKey($group.Name)) { $categoryNames[$group.Name] } else { $group.Name }
    $lines.Add("## $title")
    $lines.Add('')
    $lines.Add('| Command | Synopsis |')
    $lines.Add('|---|---|')
    foreach ($entry in ($group.Group | Sort-Object -Property Name)) {
        $lines.Add("| [$($entry.Name)]($($entry.Category)/$($entry.Name).md) | $($entry.Synopsis) |")
    }
    $lines.Add('')
}

Set-Content -Path (Join-Path $OutputPath 'README.md') -Value (($lines -join "`n").TrimEnd() + "`n") -NoNewline -Encoding utf8NoBOM

# --- Summary ----------------------------------------------------------------

foreach ($name in $incomplete) {
    Write-Warning "$name has incomplete help (PlatyPS placeholders in its page). Fill in its comment-based help."
}

Write-Host "Generated $($entries.Count) command pages in '$OutputPath'."
