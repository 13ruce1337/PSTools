# Disable-M365User

## SYNOPSIS

Deprovisions a user in Microsoft 365 (Exchange Online + Entra ID).

## SYNTAX

### __AllParameterSets

```
Disable-M365User [-UserPrincipalName] <string> [-ManagerUserPrincipalName] <string>
 [[-Date] <string>] [[-ExcludeGroup] <string[]>]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

- Sets an out-of-office reply pointing to the manager
- Converts the mailbox to a shared mailbox
- Grants the manager FullAccess to the mailbox
- Removes the user from cloud distribution lists
- Removes the user from cloud-managed Entra/M365 groups
- Removes mobile device partnerships
- Removes directly assigned licenses

Groups synced from on-prem AD are skipped (handle them with Disable-ADUser).
Dynamic groups and group-based licenses cannot be removed directly and are skipped/warned.

Requires a Microsoft Graph session with the delegated scopes:
    User.ReadWrite.All, Organization.Read.All, Group.Read.All, GroupMember.ReadWrite.All
An existing Graph session with these scopes is reused; otherwise you will be prompted to sign in.

An existing Exchange Online session is reused; if there is none, you will be prompted to sign in.

The signed-in account for both Microsoft Graph and Exchange Online is displayed
so it is clear which identity (and tenant) the changes will be made as.

## EXAMPLES

### EXAMPLE 1

Disable-M365User -UserPrincipalName trevor.cooper@contoso.com -ManagerUserPrincipalName tuan.pham@contoso.com -Date "10/10/24"

## PARAMETERS

### -Date

Date to show in the out-of-office message.
Defaults to today.

```yaml
Type: System.String
DefaultValue: (Get-Date).ToShortDateString()
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 2
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -ExcludeGroup

Display names of groups to never remove the user from.

```yaml
Type: System.String[]
DefaultValue: "@('All Users')"
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 3
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -ManagerUserPrincipalName

UPN of the manager who will receive mailbox access and be named in the out-of-office message.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
- ManagerUPN
ParameterSets:
- Name: (All)
  Position: 1
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -UserPrincipalName

UPN of the user being deprovisioned.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
- UPN
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String

{{ Fill in the Description }}

## OUTPUTS

## NOTES

Requires the Microsoft.Graph.Authentication, Microsoft.Graph.Users,
Microsoft.Graph.Users.Actions, Microsoft.Graph.Groups and
ExchangeOnlineManagement modules.
If any are missing you are asked
whether to install them.


## RELATED LINKS

{{ Fill in the related links here }}

