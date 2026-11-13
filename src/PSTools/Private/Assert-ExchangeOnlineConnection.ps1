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
        into a session is used. Graph typically ships the newer MSAL, which Exchange
        Online can also use, but not the other way round. To avoid "Method not found
        ... WithLogging" errors from Connect-MgGraph later in the session, this
        function imports Microsoft.Graph.Authentication (if installed) before
        importing ExchangeOnlineManagement.

    .PARAMETER UserPrincipalName
        Optional UPN to use if a new connection must be established.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName
    )

    # Offers to install the module if it is missing; throws if declined or the install fails.
    Assert-Module -Name ExchangeOnlineManagement

    if (-not (Get-Module -Name ExchangeOnlineManagement))
    {
        # Load Graph's MSAL first (best effort) so later Graph connections in this session still work.
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

    Write-Verbose 'Connecting to Exchange Online.'
    Connect-ExchangeOnline @connectParams

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
