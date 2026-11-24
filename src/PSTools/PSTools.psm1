#Requires -Version 5.1

# Root module: dot-source every Private and Public function file, then export
# only the Public ones.
#
# Layout: functions are grouped by area in subfolders, for example:
#   Public\ActiveDirectory\  On-premises Active Directory
#   Public\FileSystem\       File system / disk usage
#   Public\M365\             Microsoft 365 / Exchange Online
#   Public\HyperV\           Hyper-V provisioning
#   Public\Network\          Networking (Meraki, etc.)
#
# To add a command, drop a file named after the function into the appropriate
# Public\<Area>\ folder and add its name to FunctionsToExport in PSTools.psd1.
# Each file must contain only function definitions; any top-level code runs at
# import time.
#
# Docs: docs\<Area>\<Command>.md is generated from the function's comment-based
# help by tools\Build-Docs.ps1 (run by the Docs GitHub Action on push to main).
# Write complete help (.SYNOPSIS, .DESCRIPTION, .EXAMPLE, and Graph permissions
# in .NOTES) and do not edit docs\ by hand. A new Public\<Area>\ folder becomes
# a new docs area automatically; add it to the Areas list in README.MD.

$privatePath = Join-Path $PSScriptRoot 'Private'
$publicPath  = Join-Path $PSScriptRoot 'Public'

$private = @(Get-ChildItem -Path $privatePath -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue |
    Sort-Object -Property FullName)
$public  = @(Get-ChildItem -Path $publicPath  -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue |
    Sort-Object -Property FullName)

# Private first so public functions can rely on the helpers.
foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    }
    catch {
        throw "Failed to import '$($file.FullName)': $($_.Exception.Message)"
    }
}

if ($public.Count -gt 0) {
    Export-ModuleMember -Function $public.BaseName
}
