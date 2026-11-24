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
    files in it; change the function's comment-based help instead. Pages are
    built in a staging folder first, so docs\ is only replaced if every page
    builds.

    The YAML front matter PlatyPS adds to each page is removed. It is only
    needed to convert the Markdown back into MAML help, which this repo does
    not do, and GitHub shows it as a table at the top of every page.

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

# $platyOut receives PlatyPS's raw output; $buildOut is the finished docs tree.
# docs\ is only replaced once every page has been built.
$staging  = Join-Path ([IO.Path]::GetTempPath()) ("PSTools-docs-" + [guid]::NewGuid().ToString('N'))
$platyOut = Join-Path $staging 'platyps'
$buildOut = Join-Path $staging 'docs'
$null = New-Item -Path $platyOut -ItemType Directory
$null = New-Item -Path $buildOut -ItemType Directory

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

        # PlatyPS reads help properties that may not exist (for example
        # 'description' on a parameter without .PARAMETER help) and fails under
        # strict mode. Run it in a child scope with strict mode off.
        $null = & {
            Set-StrictMode -Off
            New-MarkdownCommandHelp -CommandInfo $command -OutputFolder $platyOut -Force
        }

        $generated = Get-ChildItem -Path $platyOut -Filter "$name.md" -File -Recurse | Select-Object -First 1
        if (-not $generated) {
            throw "PlatyPS did not produce a page for '$name'."
        }

        # Normalize line endings, then drop the YAML front matter (PlatyPS
        # metadata, including ms.date) so the page starts at its heading and
        # unchanged help produces no diff.
        $text = (Get-Content -Path $generated.FullName -Raw) -replace "`r`n", "`n"
        $text = $text -replace '(?s)\A---\n.*?\n---\n\s*', ''

        if ($text -match '\{\{ Fill') {
            $incomplete.Add($name)
        }

        $categoryPath = Join-Path $buildOut $category
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

    # --- Index --------------------------------------------------------------

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

    Set-Content -Path (Join-Path $buildOut 'README.md') -Value (($lines -join "`n").TrimEnd() + "`n") -NoNewline -Encoding utf8NoBOM

    # --- Replace docs\ ------------------------------------------------------

    if (Test-Path -Path $OutputPath) {
        Remove-Item -Path $OutputPath -Recurse -Force
    }
    $null = New-Item -Path $OutputPath -ItemType Directory
    Copy-Item -Path (Join-Path $buildOut '*') -Destination $OutputPath -Recurse -Force
}
finally {
    Remove-Item -Path $staging -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Summary ----------------------------------------------------------------

foreach ($name in $incomplete) {
    Write-Warning "$name has incomplete help (PlatyPS placeholders in its page). Fill in its comment-based help."
}

Write-Host "Generated $($entries.Count) command pages in '$OutputPath'."
