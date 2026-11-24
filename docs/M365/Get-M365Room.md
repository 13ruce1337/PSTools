# Get-M365Room

## SYNOPSIS

Searches for Microsoft 365 room mailboxes.

## SYNTAX

### __AllParameterSets

```
Get-M365Room [[-RoomList] <string[]>] [-Name <string>] [-UserPrincipalName <string>]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

With no -RoomList, returns all room mailboxes in the tenant.
With -RoomList
(or piped output from Get-M365RoomList), returns the rooms that belong to
those room lists.

## EXAMPLES

### EXAMPLE 1

Get-M365Room -Name '*Conf*'

### EXAMPLE 2

Get-M365Room -RoomList 'Dallas Office'

### EXAMPLE 3

Get-M365RoomList | Out-GridView -PassThru | Get-M365Room

## PARAMETERS

### -Name

Wildcard filter matched against the room Name or DisplayName.

```yaml
Type: System.String
DefaultValue: '*'
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -RoomList

One or more room list identities (name or SMTP address).
Accepts pipeline
input from Get-M365RoomList.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases:
- PrimarySmtpAddress
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -UserPrincipalName

Optional UPN used if a new Exchange Online connection is required.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
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

### System.String[]

{{ Fill in the Description }}

## OUTPUTS

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

