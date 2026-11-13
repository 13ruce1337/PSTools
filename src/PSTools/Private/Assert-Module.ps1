function Assert-Module
{
    <#
    .SYNOPSIS
        Ensures one or more PowerShell modules are installed, offering to install any that are missing.

    .DESCRIPTION
        Checks whether each named module is available (Get-Module -ListAvailable).
        If any are missing, they are listed and you are asked whether to install
        them from the PowerShell Gallery (CurrentUser scope by default, so no
        elevation is required). Declining stops with an error that includes the
        command to install them manually.

        Use -Force to install without prompting (for unattended use). In a
        non-interactive session without -Force, an error is thrown instead of
        prompting.

        Modules are NOT imported. PowerShell auto-loads them on first use, which
        keeps the Microsoft.Identity.Client (MSAL) load order handled by
        Assert-ExchangeOnlineConnection intact (Graph before Exchange Online).

        If any module cannot be installed, a single error listing every failure
        is thrown.

    .PARAMETER Name
        Names of the modules that must be available.

    .PARAMETER Scope
        Install scope for missing modules: CurrentUser (default) or AllUsers.
        AllUsers requires an elevated session.

    .PARAMETER NoInstall
        Do not offer to install missing modules; throw an error with install instructions instead.

    .PARAMETER Force
        Install missing modules without prompting.

    .EXAMPLE
        Assert-Module -Name 'Microsoft.Graph.Authentication', 'ExchangeOnlineManagement'

    .EXAMPLE
        Assert-Module -Name ActiveDirectory -NoInstall
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Name,

        [ValidateSet('CurrentUser', 'AllUsers')]
        [string]$Scope = 'CurrentUser',

        [switch]$NoInstall,

        [switch]$Force
    )

    $requested = @($Name | Select-Object -Unique)

    $available = @(Get-Module -ListAvailable -Name $requested -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Name -Unique)

    $missing = @($requested | Where-Object { $available -notcontains $_ })

    if ($missing.Count -eq 0)
    {
        Write-Verbose "All required modules are installed: $($requested -join ', ')"
        return
    }

    $missingList = $missing -join ', '
    $installHint = "Install-Module $missingList -Scope CurrentUser"

    if ($NoInstall)
    {
        throw "Required module(s) not installed: $missingList. Install with: $installHint"
    }

    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue))
    {
        throw "Required module(s) not installed: $missingList, and PowerShellGet (Install-Module) is not available to install them."
    }

    if ($Scope -eq 'AllUsers' -and -not (Test-IsAdmin))
    {
        throw "Installing modules for AllUsers requires an elevated session. Re-run elevated or use -Scope CurrentUser."
    }

    # Ask before installing (skipped with -Force, and under -WhatIf where nothing is installed).
    if (-not $Force -and -not $WhatIfPreference)
    {
        Write-Host "Required module(s) not installed: $missingList" -ForegroundColor Yellow

        $query   = "Install $missingList from the PowerShell Gallery (scope: $Scope)?"
        $caption = 'Install required modules'

        try
        {
            $approved = $PSCmdlet.ShouldContinue($query, $caption)
        }
        catch
        {
            # Non-interactive hosts cannot prompt.
            throw "Required module(s) not installed: $missingList, and this session cannot prompt for confirmation. Install with: $installHint"
        }

        if (-not $approved)
        {
            throw "Required module(s) not installed: $missingList. Install with: $installHint"
        }
    }

    # Windows PowerShell 5.1 may default to TLS 1.0/1.1, which the PowerShell Gallery rejects.
    if ($PSVersionTable.PSEdition -ne 'Core')
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }

    $failures = [System.Collections.Generic.List[string]]::new()

    foreach ($module in $missing)
    {
        if (-not $PSCmdlet.ShouldProcess($module, "Install module from PSGallery (scope: $Scope)"))
        {
            continue
        }

        Write-Host "Installing module '$module' ($Scope)..." -ForegroundColor Yellow

        try
        {
            Install-Module -Name $module -Repository PSGallery -Scope $Scope -Force -AllowClobber -ErrorAction Stop

            if (-not (Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue))
            {
                throw 'Install-Module completed but the module is still not available.'
            }

            Write-Host "Installed module '$module'." -ForegroundColor Green
        }
        catch
        {
            $failures.Add("  $module : $($_.Exception.Message)")
        }
    }

    if ($failures.Count -gt 0)
    {
        $lines = @(
            'One or more required modules could not be installed:'
            $failures
            ''
            "Install them manually with: Install-Module <name> -Scope CurrentUser"
        )
        throw ($lines -join [Environment]::NewLine)
    }
}
