function Repair-OfficeProduct {
    <#
    .SYNOPSIS
        Repairs Office Click-to-Run, or reinstalls specific products.

    .DESCRIPTION
        Default (Repair parameter set): runs the built-in Click-to-Run repair
        against the installed Office suite.
          QuickRepair - fast, offline, fixes most corruption (default)
          FullRepair  - online repair; effectively reinstalls Office and
                        requires internet access

        Reinstall parameter set (-Reinstall -Product ...): uninstalls and then
        installs the requested products with the Office Deployment Tool.
        Equivalent to the old 'alfred reinstall <app>'.

        Requires an elevated session.

    .PARAMETER RepairType
        QuickRepair (default) or FullRepair.

    .PARAMETER Reinstall
        Uninstall then install the products given in -Product.

    .PARAMETER Product
        Products to reinstall (with -Reinstall).

    .PARAMETER Force
        Close running Office applications instead of failing.

    .EXAMPLE
        Repair-OfficeProduct

    .EXAMPLE
        Repair-OfficeProduct -RepairType FullRepair -Force

    .EXAMPLE
        Repair-OfficeProduct -Reinstall -Product Office, Visio -Force
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Repair')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ParameterSetName = 'Repair')]
        [ValidateSet('QuickRepair', 'FullRepair')]
        [string]$RepairType = 'QuickRepair',

        [Parameter(Mandatory, ParameterSetName = 'Reinstall')]
        [switch]$Reinstall,

        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Reinstall')]
        [ValidateSet('Office', 'Visio', 'Project', 'OneNote', 'Excel', 'Outlook')]
        [string[]]$Product,

        [Parameter(ParameterSetName = 'Reinstall')]
        [ValidateSet('Current', 'CurrentPreview', 'MonthlyEnterprise', 'SemiAnnual', 'SemiAnnualPreview', 'BetaChannel')]
        [string]$Channel = 'Current',

        [Parameter(ParameterSetName = 'Reinstall')]
        [ValidateSet('64', '32')]
        [string]$Architecture = '64',

        [Parameter(ParameterSetName = 'Reinstall')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Language = @('en-us'),

        [Parameter(ParameterSetName = 'Reinstall')]
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
        # --- Reinstall: uninstall + install via ODT -------------------------
        if ($PSCmdlet.ParameterSetName -eq 'Reinstall') {
            $unique = @($names | Sort-Object -Unique)
            if ($unique.Count -eq 0) {
                return
            }
            $target = ($unique -join ', ')

            if (-not $PSCmdlet.ShouldProcess($target, 'Reinstall Office product(s)')) {
                return
            }

            if (-not (Test-IsAdmin)) {
                throw 'Repair-OfficeProduct must be run from an elevated session.'
            }

            $removed = Uninstall-OfficeProduct -Product $unique -Force:$Force -Confirm:$false

            $installParams = @{
                Product      = $unique
                Channel      = $Channel
                Architecture = $Architecture
                Language     = $Language
                Force        = $Force
                Confirm      = $false
            }
            if ($ExcludeApp) {
                $installParams['ExcludeApp'] = $ExcludeApp
            }
            $installed = Install-OfficeProduct @installParams

            return [pscustomobject]@{
                Action            = 'Reinstall'
                Product           = $unique
                ProductId         = $installed.ProductId
                UninstallExitCode = $removed.ExitCode
                InstallExitCode   = $installed.ExitCode
                LogPath           = $installed.LogPath
            }
        }

        # --- Click-to-Run repair --------------------------------------------
        $c2rConfig = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
        if (-not $c2rConfig) {
            throw 'No Click-to-Run Office installation was found on this machine.'
        }

        $platform = if ($c2rConfig.Platform)      { $c2rConfig.Platform }      else { 'x64' }
        $culture  = if ($c2rConfig.ClientCulture) { $c2rConfig.ClientCulture } else { 'en-us' }

        # Always use the native Common Files folder, even from 32-bit PowerShell.
        $commonFiles = if ($env:CommonProgramW6432) { $env:CommonProgramW6432 } else { $env:CommonProgramFiles }
        $c2rClient   = Join-Path $commonFiles 'microsoft shared\ClickToRun\OfficeClickToRun.exe'

        if (-not (Test-Path -LiteralPath $c2rClient)) {
            throw "OfficeClickToRun.exe was not found at '$c2rClient'."
        }

        $target = "Office ($($c2rConfig.ProductReleaseIds))"
        if (-not $PSCmdlet.ShouldProcess($target, "Run Click-to-Run $RepairType")) {
            return
        }

        if (-not (Test-IsAdmin)) {
            throw 'Repair-OfficeProduct must be run from an elevated session.'
        }

        Close-OfficeApplication -Force:$Force

        $repairArgs = @(
            'scenario=Repair'
            "platform=$platform"
            "culture=$culture"
            "RepairType=$RepairType"
            "forceappshutdown=$(if ($Force) { 'True' } else { 'False' })"
            'DisplayLevel=False'
        )

        Write-Verbose "Running $c2rClient $($repairArgs -join ' ')"
        $proc = Start-Process -FilePath $c2rClient -ArgumentList $repairArgs -PassThru -WindowStyle Hidden

        # Don't use Start-Process -Wait: it also waits for every Click-to-Run
        # process the repair starts, and some keep running, so it never returns.
        $null = $proc.Handle      # keeps ExitCode available in Windows PowerShell 5.1
        $proc.WaitForExit()       # waits for this process only, not its child processes

        if ($proc.ExitCode -ne 0) {
            throw "Click-to-Run $RepairType exited with code $($proc.ExitCode)."
        }

        [pscustomobject]@{
            Action    = $RepairType
            Product   = @($c2rConfig.ProductReleaseIds -split ',')
            ProductId = @($c2rConfig.ProductReleaseIds -split ',')
            ExitCode  = $proc.ExitCode
            LogPath   = $null
        }
    }
}
