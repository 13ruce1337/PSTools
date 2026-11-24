# Get-M365RoomList

## SYNOPSIS

Lists Microsoft 365 room lists (locations/buildings).

## SYNTAX

### __AllParameterSets

```
Get-M365RoomList [[-Name] <string>] [-UserPrincipalName <string>]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Returns room list distribution groups from Exchange Online.
Output can be
piped to Get-M365Room to list the rooms in each room list.

## EXAMPLES

### EXAMPLE 1

Get-M365RoomList

### EXAMPLE 2

Get-M365RoomList -Name '*Dallas*' | Get-M365Room

## PARAMETERS

### -Name

Wildcard filter matched against the room list Name or DisplayName.

```yaml
Type: System.String
DefaultValue: '*'
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
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

## OUTPUTS

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

