function Get-M365Room
{
    <#
    .SYNOPSIS
        Searches for Microsoft 365 room mailboxes.

    .DESCRIPTION
        With no -RoomList, returns all room mailboxes in the tenant. With -RoomList
        (or piped output from Get-M365RoomList), returns the rooms that belong to
        those room lists.

    .PARAMETER RoomList
        One or more room list identities (name or SMTP address). Accepts pipeline
        input from Get-M365RoomList.

    .PARAMETER Name
        Wildcard filter matched against the room Name or DisplayName.

    .PARAMETER UserPrincipalName
        Optional UPN used if a new Exchange Online connection is required.

    .EXAMPLE
        Get-M365Room -Name '*Conf*'

    .EXAMPLE
        Get-M365Room -RoomList 'Dallas Office'

    .EXAMPLE
        Get-M365RoomList | Out-GridView -PassThru | Get-M365Room
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipelineByPropertyName)]
        [Alias('PrimarySmtpAddress')]
        [string[]]$RoomList,

        [string]$Name = '*',

        [string]$UserPrincipalName
    )

    begin
    {
        Assert-ExchangeOnlineConnection -UserPrincipalName $UserPrincipalName
    }

    process
    {
        if ($RoomList)
        {
            foreach ($list in $RoomList)
            {
                try
                {
                    Get-DistributionGroupMember -Identity $list -ResultSize Unlimited -ErrorAction Stop |
                        Where-Object { $_.Name -like $Name -or $_.DisplayName -like $Name } |
                        Select-Object -Property @{ Name = 'RoomList'; Expression = { $list } },
                            Name, DisplayName, PrimarySmtpAddress
                }
                catch
                {
                    Write-Error "Unable to list members of room list '$list': $_"
                }
            }
        }
        else
        {
            Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
                Where-Object { $_.Name -like $Name -or $_.DisplayName -like $Name } |
                Select-Object -Property Name, DisplayName, PrimarySmtpAddress
        }
    }
}
