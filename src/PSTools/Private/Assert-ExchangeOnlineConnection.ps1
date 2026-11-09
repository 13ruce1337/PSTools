function Assert-ExchangeOnlineConnection
{
    <#
    .SYNOPSIS
        Ensures an active Exchange Online session exists, connecting if needed.

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
