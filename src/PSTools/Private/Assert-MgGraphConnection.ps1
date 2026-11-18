function Assert-MgGraphConnection
{
    <#
    .SYNOPSIS
        Ensures an active Microsoft Graph session exists with the required scopes, connecting if needed.

    .DESCRIPTION
        If the Microsoft.Graph.Authentication module is not installed, you are
        asked whether to install it (see Assert-Module).

        Reuses an existing Microsoft Graph session when it already has every
        requested scope (a '.ReadWrite.' scope satisfies the matching '.Read.'
        scope). App-only sessions are reused as-is. Otherwise Connect-MgGraph is
        called with the requested scopes. The signed-in account (or app, for
        app-only sessions) is always displayed so it is clear which identity
        (and tenant) commands will run as.

        Newer versions of Microsoft.Graph.Authentication sign in through the
        Windows Web Account Manager (WAM) broker by default. In some hosts
        (Windows Terminal, the VS Code terminal, elevated or remote sessions) the
        broker cannot obtain a parent window and fails with a NullReferenceException
        in RuntimeBroker. To avoid this, WAM is disabled with
        'Set-MgGraphOption -DisableLoginByWAM $true' before connecting, when the
        installed module supports it. Note that Set-MgGraphOption persists this
        setting in the Graph SDK profile. Use -UseWAM to leave the setting alone.

        Callers that need both Microsoft Graph and Exchange Online should call this
        function before Assert-ExchangeOnlineConnection: once Exchange Online has
        signed in, its Microsoft.Identity.Client (MSAL) assembly is loaded and
        Connect-MgGraph fails with "Method not found". If that happens, a
        descriptive error explaining how to resolve it is thrown instead of the
        raw message.

    .PARAMETER Scopes
        Delegated permission scopes that the session must have.

    .PARAMETER UseWAM
        Do not disable the Windows Web Account Manager (WAM) broker before
        signing in. Only affects new connections.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Scopes,

        [switch]$UseWAM
    )

    # Offers to install the module if it is missing; throws if declined or the install fails.
    Assert-Module -Name Microsoft.Graph.Authentication

    if (-not (Get-Module -Name Microsoft.Graph.Authentication))
    {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    # Builds a display string such as "user@contoso.com (tenant 1234...)" or "app 'MyApp' (tenant ...)".
    $describeContext = {
        param($Context)
        if ([string]$Context.AuthType -eq 'AppOnly')
        {
            $app = [string]$Context.AppName
            if ([string]::IsNullOrWhiteSpace($app)) { $app = [string]$Context.ClientId }
            $identity = "app '$app' (app-only)"
        }
        else
        {
            $identity = [string]$Context.Account
            if ([string]::IsNullOrWhiteSpace($identity)) { $identity = 'unknown account' }
        }

        $tenant = [string]$Context.TenantId
        if (-not [string]::IsNullOrWhiteSpace($tenant)) { $identity += " (tenant $tenant)" }
        return $identity
    }

    $context = Get-MgContext

    if ($context)
    {
        if ([string]$context.AuthType -eq 'AppOnly')
        {
            Write-Host "Microsoft Graph: using existing connection as $(& $describeContext $context)" -ForegroundColor DarkCyan
            return
        }

        $missing = @($Scopes | Where-Object {
            $readWrite = $_ -replace '\.Read\.', '.ReadWrite.'
            ($context.Scopes -notcontains $_) -and ($context.Scopes -notcontains $readWrite)
        })

        if ($missing.Count -eq 0)
        {
            Write-Host "Microsoft Graph: using existing connection as $(& $describeContext $context)" -ForegroundColor DarkCyan
            return
        }

        Write-Host "Existing Graph session ($(& $describeContext $context)) is missing scopes: $($missing -join ', '). Reconnecting." -ForegroundColor Yellow
    }

    # -DisableLoginByWAM exists on Set-MgGraphOption in versions that use WAM by default.
    # Detect it by parameter rather than version so this keeps working with future releases.
    $optionCommand      = Get-Command -Name Set-MgGraphOption -ErrorAction SilentlyContinue
    $supportsDisableWam = $optionCommand -and $optionCommand.Parameters.ContainsKey('DisableLoginByWAM')

    if ($supportsDisableWam -and -not $UseWAM)
    {
        try
        {
            Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction Stop
            Write-Verbose 'Disabled WAM for Microsoft Graph sign-in (Set-MgGraphOption -DisableLoginByWAM $true). This setting persists in the Graph SDK profile.'
        }
        catch
        {
            Write-Verbose "Could not disable WAM for Microsoft Graph sign-in: $($_.Exception.Message)"
        }
    }
    elseif ($UseWAM)
    {
        Write-Verbose 'Leaving the Microsoft Graph WAM setting unchanged (-UseWAM specified).'
    }
    else
    {
        Write-Verbose 'Installed Microsoft.Graph.Authentication does not support -DisableLoginByWAM; connecting with default settings.'
    }

    Write-Verbose 'Connecting to Microsoft Graph.'
    try
    {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
    }
    catch
    {
        $message = [string]$_.Exception.Message
        $details = [string]$_.Exception.ToString()

        $isBrokerError = ($details -match 'RuntimeBroker|BrokerExtension') -or
                         ($_.Exception -is [System.NullReferenceException] -and $details -match 'Microsoft\.Identity\.Client')

        if ($isBrokerError)
        {
            $graphModule  = Get-Module -Name Microsoft.Graph.Authentication | Select-Object -First 1
            $graphVersion = if ($graphModule) { [string]$graphModule.Version } else { 'unknown' }

            if ($UseWAM -and $supportsDisableWam)
            {
                $hint = 'Run the command again without -UseWAM to sign in without the broker.'
            }
            elseif ($supportsDisableWam)
            {
                $hint = 'Run "Set-MgGraphOption -DisableLoginByWAM $true", then start a new PowerShell session and try again.'
            }
            else
            {
                $hint = 'Update the module (Update-Module Microsoft.Graph.Authentication) and start a new PowerShell session.'
            }

            $brokerMessage = "Microsoft Graph sign-in failed in the Windows sign-in broker (WAM) with Microsoft.Graph.Authentication $graphVersion. $hint"
            throw [System.InvalidOperationException]::new($brokerMessage, $_.Exception)
        }

        $isMsalConflict = ($message -match 'Method not found') -and
            ($message -match 'Microsoft\.Identity')

        if (-not $isMsalConflict)
        {
            throw
        }

        $graphModule = Get-Module -Name Microsoft.Graph.Authentication | Select-Object -First 1
        $exoModule   = Get-Module -Name ExchangeOnlineManagement | Select-Object -First 1

        $msalVersions = @([AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } |
            ForEach-Object { [string]$_.GetName().Version } |
            Sort-Object -Unique)

        $graphVersion = if ($graphModule) { [string]$graphModule.Version } else { 'not loaded' }
        $exoVersion   = if ($exoModule)   { [string]$exoModule.Version }   else { 'not loaded' }
        $msalVersion  = if ($msalVersions.Count -gt 0) { $msalVersions -join ', ' } else { 'unknown' }

        if ($exoModule)
        {
            $cause = @(
                'Connect-MgGraph failed because Exchange Online was used earlier in this PowerShell session, and the Microsoft.Identity.Client (MSAL)'
                'assembly it loaded is not compatible with Microsoft.Graph.Authentication. .NET cannot unload an assembly once it is loaded.'
            )
            $resolve = @(
                '  1. Open a new PowerShell session and run the command again. PSTools commands connect to Microsoft Graph before Exchange Online.'
                '  2. If you connect manually, run Connect-MgGraph before Connect-ExchangeOnline.'
            )
        }
        else
        {
            $cause = @(
                'Connect-MgGraph failed because an incompatible Microsoft.Identity.Client (MSAL) assembly is already loaded in this PowerShell session,'
                'probably by another module. .NET cannot unload an assembly once it is loaded.'
            )
            $resolve = @(
                '  1. Open a new PowerShell session and run the Graph-based command before any other command that signs in to Microsoft 365.'
                '  2. If it persists, update the modules and restart PowerShell:'
                '       Update-Module Microsoft.Graph.Authentication'
            )
        }

        $lines = @(
            $cause
            ''
            "  Microsoft.Graph.Authentication : $graphVersion"
            "  ExchangeOnlineManagement       : $exoVersion"
            "  Loaded Microsoft.Identity.Client: $msalVersion"
            ''
            'To resolve:'
            $resolve
            ''
            "Original error: $message"
        )

        throw ($lines -join [Environment]::NewLine)
    }

    $context = Get-MgContext
    if ($context)
    {
        Write-Host "Microsoft Graph: connected as $(& $describeContext $context)" -ForegroundColor DarkCyan
    }
    else
    {
        Write-Warning 'Microsoft Graph: connected, but the signed-in account could not be determined.'
    }
}
