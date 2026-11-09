function Get-M365RoomList {
    <#
    .SYNOPSIS
        Lists Microsoft 365 room lists (locations/buildings).

    .DESCRIPTION
        Returns room list distribution groups from Exchange Online. Output can be
        piped to Get-M365Room to list the rooms in each room list.

    .PARAMETER Name
        Wildcard filter matched against the room list Name or DisplayName.

    .PARAMETER UserPrincipalName
        Optional UPN used if a new Exchange Online connection is required.

    .EXAMPLE
        Get-M365RoomList

    .EXAMPLE
        Get-M365RoomList -Name '*Dallas*' | Get-M365Room
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$Name = '*',

        [string]$UserPrincipalName
    )

    Assert-ExchangeOnlineConnection -UserPrincipalName $UserPrincipalName

    $properties = @('Name', 'DisplayName', 'PrimarySmtpAddress')

    $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited
    $roomLists | Where-Object { ($_.Name -like $Name) -or ($_.DisplayName -like $Name) } | Select-Object -Property $properties
}
