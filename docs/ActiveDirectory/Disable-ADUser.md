---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Disable-ADUser
---

# Disable-ADUser

## SYNOPSIS

Deprovisions a user in on-premises Active Directory.

## SYNTAX

### __AllParameterSets

```
Disable-ADUser [-Identity] <string[]> [-Date <string>] [-ExcludeGroup <string[]>]
 [-TargetOU <string>] [-ResetPassword] [-HideFromAddressList] [-Server <string>]
 [-Credential <pscredential>] [-PassThru] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

For each user:
  1.
Disables the account
  2.
Stamps the description with the disable date and operator
     (the previous description is kept after it)
  3.
Removes the user from all AD groups (except the primary group and
     any -ExcludeGroup entries)
  4.
Optionally resets the password to a random value (-ResetPassword)
  5.
Optionally hides the user from Exchange address lists
     (-HideFromAddressList; syncs to Exchange Online in hybrid setups)
  6.
Optionally moves the user to another OU (-TargetOU)

Failures removing individual groups are reported and do not stop the
remaining steps.
Use -PassThru to get a result object listing the groups
that were removed (useful for auditing or rollback).

Microsoft 365 deprovisioning is handled separately by Disable-M365User.

Requires RSAT (the ActiveDirectory module).

## EXAMPLES

### EXAMPLE 1

Disable-ADUser -Identity trevor.cooper

### EXAMPLE 2

Disable-ADUser -Identity trevor.cooper -ResetPassword -HideFromAddressList -TargetOU 'OU=Disabled Users,DC=contoso,DC=com'

### EXAMPLE 3

Import-Csv .\leavers.csv | Disable-ADUser -WhatIf

### EXAMPLE 4

$result = Disable-ADUser -Identity trevor.cooper -PassThru
$result.RemovedGroups

## PARAMETERS

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
- cf
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

### -Credential

Credential to use for the AD operations.

```yaml
Type: System.Management.Automation.PSCredential
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

### -Date

Date written to the account description.
Defaults to today.

```yaml
Type: System.String
DefaultValue: (Get-Date).ToShortDateString()
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -ExcludeGroup

Names (or SamAccountNames) of groups to never remove the user from.

```yaml
Type: System.String[]
DefaultValue: '@()'
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

### -HideFromAddressList

Set msExchHideFromAddressLists on the user.
Requires the Exchange schema
extensions in AD.

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

User to deprovision: SamAccountName, DistinguishedName, GUID or SID.
Accepts pipeline input (strings, AD user objects, or objects with a
SamAccountName or user property, e.g.
from Import-Csv).

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases:
- SamAccountName
- User
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

### -PassThru

Return a result object for each user.

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

### -ResetPassword

Reset the account password to a random 32-character value.

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

### -Server

Domain controller (or domain) to run the AD operations against.

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

### -TargetOU

Distinguished name of the OU or container to move the user into.

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

### -WhatIf

Runs the command in a mode that only reports what would happen without performing the actions.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
- wi
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

### System.String

{{ Fill in the Description }}

## OUTPUTS

## NOTES

Requires RSAT (the ActiveDirectory module).


## RELATED LINKS

{{ Fill in the related links here }}

