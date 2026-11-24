---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-M365GroupMember
---

# Get-M365GroupMember

## SYNOPSIS

Lists the user members of an Entra ID / Microsoft 365 group, optionally exporting them to CSV.

## SYNTAX

### __AllParameterSets

```
Get-M365GroupMember [-GroupName] <string[]> [-Transitive] [-Export] [-OutputPath <string>]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Looks up a group by display name or email address and returns the first
name, last name and email address of each user member.
Devices, contacts
and nested groups are skipped.
Use -Transitive to also include users in
nested groups.

If the value looks like an email address it is matched against the
group's primary mail address first, then its proxy (alias) addresses,
then its display name.
Otherwise it is matched against the display name.
Additional lookups only happen when the earlier ones find nothing, so the
common case is a single request.

Results are written to the pipeline.
Use -Export to also write them to a
CSV file in -OutputPath.

## EXAMPLES

### EXAMPLE 1

Get-M365GroupMember -GroupName 'Finance Team'

### EXAMPLE 2

Get-M365GroupMember -GroupName finance@contoso.com

### EXAMPLE 3

Get-M365GroupMember -GroupName 'All Staff' -Transitive -Export -OutputPath C:\Temp

### EXAMPLE 4

'Sales', 'marketing@contoso.com' | Get-M365GroupMember | Format-Table

## PARAMETERS

### -Export

Also export the results to a CSV file in -OutputPath.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
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

### -GroupName

Display name or email address of the group.
Accepts pipeline input.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases:
- DisplayName
- Group
- Identity
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -OutputPath

Folder for the CSV export (used with -Export).
Created if it does not exist.
Default: $env:TEMP\PSTools

```yaml
Type: System.String
DefaultValue: (Join-Path $env:TEMP 'PSTools')
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

### -Transitive

Include users who are members through nested groups.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
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

Requires the Microsoft.Graph.Authentication module and the delegated scopes:
    User.Read.All, GroupMember.Read.All
An existing Graph session with these scopes is reused.


## RELATED LINKS

{{ Fill in the related links here }}

