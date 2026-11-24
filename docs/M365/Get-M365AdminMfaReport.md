---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-M365AdminMfaReport
---

# Get-M365AdminMfaReport

## SYNOPSIS

Reports the MFA registration status of every Entra ID directory role holder.

## SYNTAX

### __AllParameterSets

```
Get-M365AdminMfaReport [[-OutputPath] <string>] [-IncludeEligible] [-NoExport]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Finds all users holding an Entra ID directory role (Global Administrator,
Exchange Administrator, etc.), including users who receive a role through
a role-assignable group, then checks their registered authentication
methods and classifies their MFA strength:

    Phishing-resistant         FIDO2, Windows Hello for Business, platform credential, certificate
    Enabled                    Microsoft Authenticator, software/hardware OATH token
    Weak (phone only)          SMS / voice only
    Disabled                   Only password / email / Temporary Access Pass
    Unknown                    Authentication methods could not be read
    External (partner tenant)  Role held by a partner (CSP DAP/GDAP) group such as
                               AdminAgents / HelpdeskAgents.
Its members live in the
                               partner's tenant, so their MFA cannot be read here.

Report objects are written to the pipeline and exported to CSV in
-OutputPath unless -NoExport is specified.

## EXAMPLES

### EXAMPLE 1

Get-M365AdminMfaReport | Format-Table -AutoSize

### EXAMPLE 2

Get-M365AdminMfaReport -IncludeEligible -OutputPath C:\Temp

### EXAMPLE 3

Get-M365AdminMfaReport -NoExport | Where-Object MfaStatus -notin 'Enabled', 'Phishing-resistant'

## PARAMETERS

### -IncludeEligible

Also include PIM eligible role assignments (requires Entra ID P2).

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

### -NoExport

Do not write a CSV file.

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

### -OutputPath

Folder for the CSV export.
Created if it does not exist.
Default: $env:TEMP\PSTools

```yaml
Type: System.String
DefaultValue: (Join-Path $env:TEMP 'PSTools')
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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

Requires the Microsoft.Graph.Authentication module and the delegated scopes:
    User.Read.All, RoleManagement.Read.Directory, GroupMember.Read.All,
    UserAuthenticationMethod.Read.All
An existing Graph session with these scopes is reused.


## RELATED LINKS

{{ Fill in the related links here }}

