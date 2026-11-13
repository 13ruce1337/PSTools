# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

The module version is defined only in `src/PSTools/PSTools.psd1` (`ModuleVersion`).
Individual functions are not versioned separately.

## [Unreleased]

### Added
- M365: `Get-M365UserSummary` - admin-center style overview of a user: account
  status, password last set, sign-in activity, authentication methods, licenses,
  groups and roles, mailbox details, mailbox delegates (FullAccess / SendAs /
  SendOnBehalf), forwarding, out-of-office and mobile devices.

### Planned
- Import remaining functions from other repositories.
- Comment-based help / documentation for every public command.
- GitHub releases and PowerShell Gallery publishing (target: 1.0.0).

## [0.6.0]

First tracked version. Earlier per-function version numbers have been retired
in favor of a single module version. History from those notes is summarized below.

### Added
- ActiveDirectory: `Disable-ADUser`.
- FileSystem: `Get-DiskUsageReport`.
- HyperV: `New-DebianVM`, `New-WindowsVM`.
- M365: `Disable-M365User`, `Get-M365AdminMfaReport`, `Get-M365GroupMember`,
  `Get-M365MailboxSpaceReport`, `Get-M365Room`, `Get-M365RoomList`.
- Network: `Get-MerakiFirmware`.
- Private helpers: `Assert-ExchangeOnlineConnection`, `Assert-MgGraphConnection`,
  `Assert-Module`, `Test-IsAdmin`.

### Changed
- Versioning reset to SemVer `0.x` with the manifest as the single source of truth
  (the manifest previously said `2.6.0`).
- `Disable-ADUser`: split from the original `Disable-User` script into an AD-only
  function; renamed from `Disable-ADUserAccount` to match its file name.
- `Disable-M365User`: split from the original `Disable-User` script into an
  M365-only function; replaced AzureAD with Microsoft Graph; uses
  `Assert-ExchangeOnlineConnection` and `Assert-MgGraphConnection`; uses
  `Invoke-MgGraphRequest` for user/group lookups (SDK-version independent,
  paged, UPN-encoded); accepts CSV/pipeline input.

### Known issues
- `Get-MerakiFirmware` uses the deprecated `/networks/{id}/devices` endpoint.
