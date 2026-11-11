function Assert-MgGraphConnection
{
    <#
    .SYNOPSIS
        Ensures an active Microsoft Graph session exists with the required scopes, connecting if needed.

    .DESCRIPTION
        Reuses an existing Microsoft Graph session when it already has every
        requested scope (a '.ReadWrite.' scope satisfies the matching '.Read.'
        scope). App-only sessions are reused as-is. Otherwise Connect-MgGraph is
        called with the requested scopes.

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

    $context = Get-MgContext

    if ($context)
    {
        if ([string]$context.AuthType -eq 'AppOnly')
        {
            Write-Verbose "Using existing app-only Microsoft Graph connection (app '$($context.AppName)')."
            return
        }

        $missing = @($Scopes | Where-Object {
            $readWrite = $_ -replace '\.Read\.', '.ReadWrite.'
            ($context.Scopes -notcontains $_) -and ($context.Scopes -notcontains $readWrite)
        })

        if ($missing.Count -eq 0)
        {
            Write-Verbose "Using existing Microsoft Graph connection as '$($context.Account)'."
            return
        }

        Write-Host "Existing Graph session is missing scopes: $($missing -join ', '). Reconnecting." -ForegroundColor Yellow
    }

    Write-Verbose 'Connecting to Microsoft Graph.'
    Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
}
