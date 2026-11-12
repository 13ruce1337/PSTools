function Assert-MgGraphConnection
{
    <#
    .SYNOPSIS
        Ensures an active Microsoft Graph session exists with the required scopes, connecting if needed.

    .DESCRIPTION
        Reuses an existing Microsoft Graph session when it already has every
        requested scope (a '.ReadWrite.' scope satisfies the matching '.Read.'
        scope). App-only sessions are reused as-is. Otherwise Connect-MgGraph is
        called with the requested scopes. The signed-in account (or app, for
        app-only sessions) is always displayed so it is clear which identity
        (and tenant) commands will run as.

        If the connection fails because an incompatible Microsoft.Identity.Client
        (MSAL) assembly was already loaded by another module (typically
        ExchangeOnlineManagement), a descriptive error explaining how to resolve
        it is thrown instead of the raw "Method not found" message.

    .PARAMETER Scopes
        Delegated permission scopes that the session must have.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Scopes
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication))
    {
        throw "The Microsoft.Graph.Authentication module is not installed. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }

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

    Write-Verbose 'Connecting to Microsoft Graph.'
    try
    {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
    }
    catch
    {
        $message = [string]$_.Exception.Message

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

        $lines = @(
            'Connect-MgGraph failed because an incompatible Microsoft.Identity.Client (MSAL) assembly is already loaded in this PowerShell session.'
            'This usually happens when ExchangeOnlineManagement was loaded before Microsoft.Graph.Authentication; .NET cannot unload the older MSAL.'
            ''
            "  Microsoft.Graph.Authentication : $graphVersion"
            "  ExchangeOnlineManagement       : $exoVersion"
            "  Loaded Microsoft.Identity.Client: $msalVersion"
            ''
            'To resolve:'
            '  1. Open a new PowerShell session and run the Graph-based command (or Connect-MgGraph) before any Exchange Online command.'
            '  2. If it persists, update both modules and restart PowerShell:'
            '       Update-Module Microsoft.Graph.Authentication, ExchangeOnlineManagement'
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
