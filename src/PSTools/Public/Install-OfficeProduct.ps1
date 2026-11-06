function Install-OfficeProduct {
    <#
    .SYNOPSIS
        Installs Microsoft 365 / Office Click-to-Run products with the Office
        Deployment Tool.

    .DESCRIPTION
        Builds a single ODT configuration for every requested product and runs
        setup.exe once. The ODT is downloaded and cached automatically.

        Product mapping:
          Office  -> O365ProPlusRetail (Microsoft 365 Apps for enterprise)
          Visio   -> VisioPro2021Retail
          Project -> ProjectPro2021Retail
          OneNote -> OneNoteRetail
          Excel   -> Excel2019Retail
          Outlook -> OutlookRetail

        Requires an elevated session.

    .PARAMETER Product
        One or more products to install. Accepts pipeline input.

    .PARAMETER Channel
        Update channel. Default 'Current'.

    .PARAMETER Architecture
        '64' (default) or '32'. Must match any Office already installed.

    .PARAMETER Language
        One or more language IDs (e.g. en-us, fr-fr, MatchOS). Default en-us.

    .PARAMETER ExcludeApp
        Apps to exclude from the installed products.

    .PARAMETER Force
        Close running Office applications instead of failing.

    .EXAMPLE
        Install-OfficeProduct Office

    .EXAMPLE
        Install-OfficeProduct Office, Visio, Project -ExcludeApp Publisher, Access

    .EXAMPLE
        'Visio', 'Project' | Install-OfficeProduct -Force -Verbose

    .EXAMPLE
        Install-OfficeProduct Office -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateSet('Office', 'Visio', 'Project', 'OneNote', 'Excel', 'Outlook')]
        [string[]]$Product,

        [ValidateSet('Current', 'CurrentPreview', 'MonthlyEnterprise', 'SemiAnnual', 'SemiAnnualPreview', 'BetaChannel')]
        [string]$Channel = 'Current',

        [ValidateSet('64', '32')]
        [string]$Architecture = '64',

        [ValidateNotNullOrEmpty()]
        [string[]]$Language = @('en-us'),

        [ValidateSet('Access', 'Bing', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word')]
        [string[]]$ExcludeApp,

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
        $unique = @($names | Sort-Object -Unique)
        if ($unique.Count -eq 0) {
            return
        }

        $ids    = @($unique | Get-OfficeProductId)
        $target = ($unique -join ', ')

        if (-not $PSCmdlet.ShouldProcess($target, 'Install Office product(s)')) {
            return
        }

        if (-not (Test-IsAdmin)) {
            throw 'Install-OfficeProduct must be run from an elevated session.'
        }

        $logPath = Join-Path $env:TEMP 'PSTools\ODT\Logs'

        $configParams = @{
            Action       = 'Add'
            ProductId    = $ids
            Channel      = $Channel
            Architecture = $Architecture
            Language     = $Language
            LogPath      = $logPath
        }
        if ($ExcludeApp) {
            $configParams['ExcludeApp'] = $ExcludeApp
        }

        $xml = New-OdtConfiguration @configParams

        Write-Verbose "Installing: $target"
        $exitCode = Invoke-OdtSetup -Configuration $xml -LogPath $logPath -Force:$Force

        [pscustomobject]@{
            Action    = 'Install'
            Product   = $unique
            ProductId = $ids
            ExitCode  = $exitCode
            LogPath   = $logPath
        }
    }
}
