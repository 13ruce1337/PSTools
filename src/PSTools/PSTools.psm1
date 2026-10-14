#Requires -Version 5.1

# Root module: dot-source every Private and Public function file, then export
# only the Public ones. To add a command, drop a file named after the function
# into Public/ and add its name to FunctionsToExport in PSTools.psd1.

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
$public  = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public')  -Filter '*.ps1' -File -ErrorAction SilentlyContinue)

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
