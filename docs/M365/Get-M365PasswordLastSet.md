---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-M365PasswordLastSet
---

# Get-M365PasswordLastSet

## SYNOPSIS

Shows the Microsoft 365 users whose password was set longest ago.

## SYNTAX

### __AllParameterSets

```
Get-M365PasswordLastSet [[-Identity] <string[]>] [-Top <int>] [-EnabledOnly] [-Export]
 [-OutputPath <string>] [-PassThru]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Reads the password last set date of every user in the tenant (or the
users given with -Identity) from Microsoft Graph.
It then prints the
-Top users (default 30) with the oldest password, longest first.
Users with no password last set date are listed last.

Nothing is written to disk unless -Export or -OutputPath is specified.
When exporting, the full sorted list (not only the top entries) is
written to a CSV file in -OutputPath.

Use -PassThru to also return every record to the pipeline.

## EXAMPLES

### EXAMPLE 1

Get-M365PasswordLastSet

### EXAMPLE 2

Get-M365PasswordLastSet -Top 50 -EnabledOnly

### EXAMPLE 3

Get-M365PasswordLastSet -OutputPath C:\Reports

### EXAMPLE 4

Get-M365PasswordLastSet -Identity user@contoso.com

### EXAMPLE 5

Get-M365PasswordLastSet -PassThru | Where-Object { $_.IsActive -and $_.PasswordAgeDays -gt 365 }

## PARAMETERS

### -EnabledOnly

Only include enabled accounts.

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

### -Export

Export the full results to a CSV file in -OutputPath.

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

### -Identity

Optional.
Specific users to report on: UPN, object ID or email address
(primary or alias).
Accepts pipeline input.
If omitted, all users in the
tenant are read.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases:
- UserPrincipalName
- UPN
- PrimarySmtpAddress
- Mail
- Id
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: false
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -OutputPath

Folder for the CSV export.
Created if it does not exist.
Specifying this
parameter also turns on the export.
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

### -PassThru

Return all records to the pipeline, in addition to the on-screen report.

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

### -Top

Number of users to show on screen.
Default 30.

```yaml
Type: System.Int32
DefaultValue: 30
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

Requires the Microsoft.Graph.Authentication module and the delegated scope:
    User.Read.All
An existing Graph session with this scope is reused.

For users synced from on-premises AD, the password last set date reflects
the last password change synced by Entra Connect (password hash sync).


## RELATED LINKS

{{ Fill in the related links here }}

