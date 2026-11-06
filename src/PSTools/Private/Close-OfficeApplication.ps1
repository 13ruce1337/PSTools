function Close-OfficeApplication {
    <#
    .SYNOPSIS
        Ensures no Office applications are running.

    .DESCRIPTION
        Finds processes running from a 'Microsoft Office' install folder.
        Without -Force, throws if any are found so the user can save their
        work. With -Force, terminates them.

        The Click-to-Run service (OfficeClickToRun.exe) lives under
        'Common Files\microsoft shared\ClickToRun' and is not affected.
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    $running = @(
        Get-CimInstance -ClassName Win32_Process `
            -Filter "ExecutablePath LIKE '%Microsoft Office%'" `
            -ErrorAction SilentlyContinue
    )

    if ($running.Count -eq 0) {
        Write-Verbose 'No Office applications are running.'
        return
    }

    $names = ($running.Name | Sort-Object -Unique) -join ', '

    if (-not $Force) {
        throw "Office applications are running ($names). Close them and try again, or use -Force to close them automatically."
    }

    Write-Verbose "Closing Office applications: $names"
    foreach ($proc in $running) {
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
    }
}
