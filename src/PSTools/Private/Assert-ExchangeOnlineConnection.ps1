function Assert-ExchangeOnlineConnection
{
    <#
    .SYNOPSIS
        Ensures an active Exchange Online session exists, connecting if needed.

    .DESCRIPTION
        Reuses an existing Exchange Online session if one is connected, otherwise
        calls Connect-ExchangeOnline.

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

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement))
    {
        throw "The ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

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

    $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' } |
        Select-Object -First 1

    if ($connection)
    {
        Write-Verbose "Using existing Exchange Online connection as '$($connection.UserPrincipalName)'."
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
}
