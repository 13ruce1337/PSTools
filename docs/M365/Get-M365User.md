---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-M365User
---

# Get-M365User

## SYNOPSIS

Shows a brief overview of a Microsoft 365 user, similar to the user page
in the Microsoft 365 admin center.

## SYNTAX

### __AllParameterSets

```
Get-M365User [-Identity] <string[]> [-SkipMailbox] [-PassThru]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Collects the key facts about a user from Microsoft Graph and Exchange Online:

    Account         Enabled / blocked, user type, cloud or synced, created,
                    job title, department, office, phone, manager
    Password        Last set, age in days, never-expires policy
    Sign-in         Last interactive / non-interactive / successful sign-in
                    (requires Entra ID P1 and AuditLog.Read.All)
    Auth methods    Registered authentication (MFA) methods
    Licenses        Assigned SKUs, direct or group-based, assignment errors
    Groups / roles  Group memberships and activated directory roles
    Mailbox         Type, size, item count, last logon, archive, litigation
                    hold, hidden from GAL, forwarding
    Mailbox access  FullAccess, SendAs and SendOnBehalf delegates
    Out of office   State, schedule, audience and messages
    Mobile devices  Devices with an Exchange ActiveSync partnership

Each section is retrieved independently; if one cannot be read a note is
shown and the rest of the report continues.
Use -Verbose to see the full
error details.

The report is written to the console.
Use -PassThru to also return a
report object.

## EXAMPLES

### EXAMPLE 1

Get-M365User -Identity trevor.cooper@contoso.com

### EXAMPLE 2

Get-M365User trevor.cooper@contoso.com -SkipMailbox

### EXAMPLE 3

$summary = Get-M365User -Identity trevor.cooper@contoso.com -PassThru
$summary.MailboxPermissions

### EXAMPLE 4

'user1@contoso.com', 'user2@contoso.com' | Get-M365User

## PARAMETERS

### -Identity

The user to look up: UPN, object ID or email address (primary or alias).
Accepts pipeline input.

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
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -PassThru

Return a report object in addition to the on-screen report.

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

### -SkipMailbox

Skip the Exchange Online sections (no Exchange Online connection is made).

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
    User.Read.All, GroupMember.Read.All, UserAuthenticationMethod.Read.All,
    AuditLog.Read.All
Mailbox sections require the ExchangeOnlineManagement module.
Existing Graph and Exchange Online sessions are reused.

For users synced from on-premises AD, the password last set date reflects
the last password change synced by Entra Connect (password hash sync).


## RELATED LINKS

{{ Fill in the related links here }}

