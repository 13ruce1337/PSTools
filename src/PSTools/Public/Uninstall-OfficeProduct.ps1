function Uninstall-OfficeProduct {
    <#
    .SYNOPSIS
        Removes Office Click-to-Run products with the Office Deployment Tool.

    .DESCRIPTION
        Removes the requested products in a single ODT run, or every
        Click-to-Run Office product with -All (the old 'alfred nuke office').

        Requires an elevated session.

    .PARAMETER Product
        One or more products to remove. Accepts pipeline input.

    .PARAMETER All
        Remove every Click-to-Run Office product on the machine.

    .PARAMETER Force
        Close running Office applications instead of failing.

    .EXAMPLE
        Uninstall-OfficeProduct Visio

    .EXAMPLE
        Uninstall-OfficeProduct Visio, Project -Force -Confirm:$false

    .EXAMPLE
        Uninstall-OfficeProduct -All -Force
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'ByProduct')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ByProduct')]
        [ValidateSet('Office', 'Visio', 'Project', 'OneNote', 'Excel', 'Outlook')]
        [string[]]$Product,

        [Parameter(Mandatory, ParameterSetName = 'All')]
        [switch]$All,

        [switch]$Force
    )

    begin {
        $names = New-Object System.Collections.Generic.List[string]
    }

    process {
        foreach ($p in $Product) {
            $names.Add($p)
        }
    }

    end {
        if ($PSCmdlet.ParameterSetName -eq 'All') {
            $unique = @('All')
            $ids    = @()
            $target = 'All Click-to-Run Office products'
        }
        else {
            $unique = @($names | Sort-Object -Unique)
            if ($unique.Count -eq 0) {
                return
            }
            $ids    = @($unique | Get-OfficeProductId)
            $target = ($unique -join ', ')
        }

        if (-not $PSCmdlet.ShouldProcess($target, 'Uninstall Office product(s)')) {
            return
        }

        if (-not (Test-IsAdmin)) {
            throw 'Uninstall-OfficeProduct must be run from an elevated session.'
        }

        $logPath = Join-Path $env:TEMP 'PSTools\ODT\Logs'

        if ($PSCmdlet.ParameterSetName -eq 'All') {
            $xml = New-OdtConfiguration -Action RemoveAll -LogPath $logPath
        }
        else {
            $xml = New-OdtConfiguration -Action Remove -ProductId $ids -LogPath $logPath
        }

        Write-Verbose "Uninstalling: $target"
        $exitCode = Invoke-OdtSetup -Configuration $xml -LogPath $logPath -Force:$Force

        [pscustomobject]@{
            Action    = 'Uninstall'
            Product   = $unique
            ProductId = $ids
            ExitCode  = $exitCode
            LogPath   = $logPath
        }
    }
}
