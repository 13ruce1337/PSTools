function Assert-ExchangeOnlineConnection
{
    <#
    .SYNOPSIS
        Ensures an active Exchange Online session exists, connecting if needed.

    .DESCRIPTION
        If the ExchangeOnlineManagement module is not installed, you are asked
        whether to install it (see Assert-Module).

        Reuses an existing Exchange Online session if one is connected, otherwise
        calls Connect-ExchangeOnline. The signed-in account is always displayed so
        it is clear which identity (and tenant) commands will run as.

        ExchangeOnlineManagement and Microsoft.Graph.Authentication each ship their
        own Microsoft.Identity.Client (MSAL) assembly, and only the first one loaded
        into a session is used. Once Connect-ExchangeOnline has signed in, its MSAL
        is loaded and a later Connect-MgGraph fails with "Method not found".
        Callers that need both services should call Assert-MgGraphConnection
        before this function. As a best effort, this function also imports
        Microsoft.Graph.Authentication (if installed) before importing
        ExchangeOnlineManagement.

        From ExchangeOnlineManagement 3.7.0, Connect-ExchangeOnline signs in through
        the Windows Web Account Manager (WAM) broker by default. In some hosts
        (Windows Terminal, the VS Code terminal, elevated or remote sessions) the
        broker cannot obtain a parent window and fails with a NullReferenceException
        in RuntimeBroker. To avoid this, WAM is disabled (-DisableWAM) whenever the
        installed module supports it (3.7.2 and later). Use -UseWAM to opt back in.

        If sign-in still fails in the broker, a short, actionable error is thrown
        instead of the raw MSAL stack trace. The original exception is kept as the
        inner exception.

    .PARAMETER UserPrincipalName
        Optional UPN to use if a new connection must be established.

    .PARAMETER UseWAM
        Use the Windows Web Account Manager (WAM) broker for sign-in instead of
        disabling it. Only affects new connections.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,

        [switch]$UseWAM
    )

    # Offers to install the module if it is missing; throws if declined or the install fails.
    Assert-Module -Name ExchangeOnlineManagement

    if (-not (Get-Module -Name ExchangeOnlineManagement))
    {
        # Load Graph first (best effort) so later Graph connections in this session are less likely to conflict.
        if (-not (Get-Module -Name Microsoft.Graph.Authentication) -and
            (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication))
        {
            try
            {
                Write-Verbose 'Pre-loading Microsoft.Graph.Authentication to avoid MSAL assembly conflicts with ExchangeOnlineManagement.'
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            }
            catch
            {
                Write-Verbose "Could not pre-load Microsoft.Graph.Authentication: $($_.Exception.Message)"
            }
        }

        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    }

    # Returns the first connected Exchange Online session, or $null.
    $getConnection = {
        Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' } |
            Select-Object -First 1
    }

    # Builds a display string such as "user@contoso.com (tenant 1234...)".
    $describeConnection = {
        param($Connection)
        $account = [string]$Connection.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($account)) { $account = 'unknown account' }
        $tenant = [string]$Connection.TenantID
        if (-not [string]::IsNullOrWhiteSpace($tenant)) { $account += " (tenant $tenant)" }
        return $account
    }

    # True when an error came from the WAM broker (RuntimeBroker) in MSAL.
    $testBrokerError = {
        param($ErrorRecord)
        $details = [string]$ErrorRecord.Exception.ToString()
        return ($details -match 'RuntimeBroker|BrokerExtension') -or
               ($ErrorRecord.Exception -is [System.NullReferenceException] -and $details -match 'Microsoft\.Identity\.Client')
    }

    $connection = & $getConnection

    if ($connection)
    {
        Write-Host "Exchange Online: using existing connection as $(& $describeConnection $connection)" -ForegroundColor DarkCyan
        return
    }

    $connectParams = @{
        ShowBanner  = $false
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName))
    {
        $connectParams.UserPrincipalName = $UserPrincipalName
    }

    $exoModule  = Get-Module -Name ExchangeOnlineManagement | Select-Object -First 1
    $exoVersion = if ($exoModule) { [string]$exoModule.Version } else { 'unknown' }

    # -DisableWAM exists in ExchangeOnlineManagement 3.7.2 and later. Detect it by
    # parameter rather than version so this keeps working with future releases.
    $connectCommand     = Get-Command -Name Connect-ExchangeOnline -ErrorAction Stop
    $supportsDisableWam = $connectCommand.Parameters.ContainsKey('DisableWAM')

    if ($supportsDisableWam -and -not $UseWAM)
    {
        $connectParams.DisableWAM = $true
        Write-Verbose "Connecting to Exchange Online (ExchangeOnlineManagement $exoVersion) with WAM disabled (-DisableWAM)."
    }
    elseif ($UseWAM)
    {
        Write-Verbose "Connecting to Exchange Online (ExchangeOnlineManagement $exoVersion) using the WAM broker (-UseWAM specified)."
    }
    else
    {
        Write-Verbose "Connecting to Exchange Online (ExchangeOnlineManagement $exoVersion does not support -DisableWAM)."
    }

    try
    {
        Connect-ExchangeOnline @connectParams
    }
    catch
    {
        if (-not (& $testBrokerError $_))
        {
            throw
        }

        if ($UseWAM -and $supportsDisableWam)
        {
            $hint = 'Run the command again without -UseWAM to sign in without the broker.'
        }
        elseif (-not $supportsDisableWam)
        {
            $hint = 'This version does not support -DisableWAM. Update the module (Update-Module ExchangeOnlineManagement) to 3.7.2 or later ' +
                    'and start a new PowerShell session.'
        }
        else
        {
            $hint = 'Sign-in failed even with WAM disabled. Open a new PowerShell session and try again. ' +
                    'If that also fails, try a classic PowerShell console window instead of Windows Terminal or VS Code.'
        }

        $message = "Exchange Online sign-in failed in the Windows sign-in broker (WAM) with ExchangeOnlineManagement $exoVersion. $hint"
        throw [System.InvalidOperationException]::new($message, $_.Exception)
    }

    $connection = & $getConnection

    if ($connection)
    {
        Write-Host "Exchange Online: connected as $(& $describeConnection $connection)" -ForegroundColor DarkCyan
    }
    else
    {
        Write-Warning 'Exchange Online: connected, but the signed-in account could not be determined.'
    }
}
