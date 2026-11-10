function Disable-User {
    <#
    .SYNOPSIS
        Disables/Deprovisions a user in a hybrid AD + Microsoft 365 environment.

    .DESCRIPTION
        Wrapper that connects to services once, then for each user runs:
            1. Disable-ADUserAccount  (on-prem AD)
            2. Disable-M365User       (Exchange Online / Entra ID)

        Run with -User/-Manager for a single user, or with no arguments (or -CsvPath)
        to process a CSV with the columns: user, manager, date

        Existing Microsoft Graph / Exchange Online sessions are reused and left connected.
        Sessions opened by this function are disconnected when it finishes.

        IMPORTANT: Requires RSAT (ActiveDirectory module). The AD lookup is always used to
        resolve UPNs, even with -SkipAD.

    .PARAMETER User
        SamAccountName of the user to deprovision.

    .PARAMETER Manager
        SamAccountName of the user's manager.

    .PARAMETER Date
        Date to show in the out-of-office message. Defaults to today.

    .PARAMETER CsvPath
        Path to a CSV with columns user, manager, date. If omitted (and -User is not used),
        a file picker is shown.

    .PARAMETER SkipAD
        Skip the on-prem AD deprovisioning step.

    .PARAMETER SkipM365
        Skip the Microsoft 365 deprovisioning step (and skip connecting to M365).

    .NOTES
        Version:        2.1
        Author:         Trevor Cooper
        Creation Date:  10/17/24
        Purpose/Change: Initial script development
        Mod 5/7/25:  Updated QOL, comments, function for connecting, module check/install
        Mod 5/14/25: Added function for main components, added ability to take a CSV
        Mod 2.0:     Split into Disable-ADUserAccount and Disable-M365User; this is now a wrapper
        Mod 2.1:     Use Assert-ExchangeOnlineConnection; only disconnect sessions this function opened

    .EXAMPLE
        Disable-User -User trevor.cooper -Manager tuan.pham -Date "10/10/24"

    .EXAMPLE
        Disable-User

    .EXAMPLE
        Disable-User -CsvPath .\leavers.csv -SkipAD
    #>
    [CmdletBinding(DefaultParameterSetName = 'Csv')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Single')]
        [string]$User,

        [Parameter(Mandatory, ParameterSetName = 'Single')]
        [string]$Manager,

        [Parameter(ParameterSetName = 'Single')]
        [string]$Date = (Get-Date).ToShortDateString(),

        [Parameter(ParameterSetName = 'Csv')]
        [string]$CsvPath,

        [switch]$SkipAD,

        [switch]$SkipM365
    )

    $ProgressPreference = 'SilentlyContinue'

    $graphScopes = @(
        'User.ReadWrite.All',
        'Organization.Read.All',
        'Group.Read.All',
        'GroupMember.ReadWrite.All'
    )

    $logFile = Join-Path -Path $env:TEMP -ChildPath 'deprovision_user.log'
    Start-Transcript -Force -IncludeInvocationHeader -Path $logFile | Out-Null

    # Track which sessions this function opened so we only tear down our own.
    $openedGraph = $false
    $openedExo   = $false

    try {
        #------------------------------------------------------------[Build user list]------------------------------------------------------------

        if ($PSCmdlet.ParameterSetName -eq 'Single') {
            Write-Host "Running for a single user"
            $userList = @([pscustomobject]@{ user = $User; manager = $Manager; date = $Date })
        }
        else {
            Write-Host "Running for multiple users"
            if (-not $CsvPath) {
                Add-Type -AssemblyName System.Windows.Forms
                $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $fileDialog.Filter = 'CSV (*.csv)|*.csv'
                if ($fileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                    throw "No CSV file selected."
                }
                $CsvPath = $fileDialog.FileName
            }
            $userList = @(Import-Csv -Path $CsvPath -ErrorAction Stop)
        }

        #------------------------------------------------------------[Dependencies]------------------------------------------------------------

        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            throw "The ActiveDirectory module was not found. Install RSAT (Rsat.ActiveDirectory.DS-LDS.Tools) to use this function."
        }
        Import-Module ActiveDirectory -ErrorAction Stop

        if (-not $SkipM365) {
            Assert-Module -Name @(
                'Microsoft.Graph.Authentication',
                'Microsoft.Graph.Users',
                'Microsoft.Graph.Users.Actions',
                'Microsoft.Graph.Groups',
                'ExchangeOnlineManagement'
            )

            Write-Host "Due to MFA, we cannot store credentials. You may be asked to sign in for each service." -ForegroundColor Yellow

            # Microsoft Graph: reuse an existing session only if it has every scope we need
            $mgContext = Get-MgContext
            $missingScopes = @()
            if ($mgContext) {
                $missingScopes = @($graphScopes | Where-Object { $mgContext.Scopes -notcontains $_ })
            }
            if (-not $mgContext -or $missingScopes.Count -gt 0) {
                if ($mgContext) {
                    Write-Host "Existing Graph session is missing scopes: $($missingScopes -join ', '). Reconnecting." -ForegroundColor Yellow
                }
                else {
                    $openedGraph = $true
                }
                Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
            }
            else {
                Write-Verbose "Using existing Microsoft Graph connection as '$($mgContext.Account)'."
            }

            # Exchange Online: reuse an existing session, otherwise connect via the shared helper
            $exoPreconnected = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' })
            Assert-ExchangeOnlineConnection
            $openedExo = -not $exoPreconnected
        }

        #------------------------------------------------------------[Execution]------------------------------------------------------------

        foreach ($emp in $userList) {
            try {
                $adUser    = Get-ADUser -Identity $emp.user -Properties UserPrincipalName -ErrorAction Stop
                $adManager = Get-ADUser -Identity $emp.manager -Properties UserPrincipalName -ErrorAction Stop
                $empDate   = if ($emp.date) { $emp.date } else { (Get-Date).ToShortDateString() }

                Write-Host "===== Deprovisioning $($adUser.SamAccountName) =====" -ForegroundColor Green

                if (-not $SkipAD) {
                    Disable-ADUserAccount -Identity $adUser.SamAccountName
                }

                if (-not $SkipM365) {
                    Disable-M365User -UserPrincipalName $adUser.UserPrincipalName -ManagerUserPrincipalName $adManager.UserPrincipalName -Date $empDate
                }
            }
            catch {
                Write-Warning "Failed to deprovision $($emp.user): $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Error $_
    }
    finally {
        if ($openedExo) {
            Write-Host "Disconnecting from Exchange Online"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        if ($openedGraph) {
            Write-Host "Disconnecting from Microsoft Graph"
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        Write-Host "Transcript can be found: $logFile"
        Stop-Transcript | Out-Null
    }
}
