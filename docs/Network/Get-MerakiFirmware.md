---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: https://developer.cisco.com/meraki/api-v1/
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-MerakiFirmware
---

# Get-MerakiFirmware

## SYNOPSIS

Get Meraki Firmware

## SYNTAX

### __AllParameterSets

```
Get-MerakiFirmware [-api_key] <string>
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Using Meraki's API, it grabs all devices and lists some information including the firmware.

It also grabs the latest firmware version.
!Remember to enter your API key!

## EXAMPLES

## PARAMETERS

### -api_key

{{ Fill api_key Description }}

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
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

### Outputs a CSV file in the same directory as this file.

{{ Fill in the Description }}

## NOTES

Uses the deprecated /networks/$($network.id)/devices endpoint; this is due to be replaced.


## RELATED LINKS

- [](https://developer.cisco.com/meraki/api-v1/)
