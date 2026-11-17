# Changelog

## [0.6.0]

First tracked version. Earlier per-function version numbers have been retired
in favor of a single module version. History from those notes is summarized below.

### Added
- M365: `Get-M365PasswordLastSet`
- M365: `Get-M365UserSummary`
- ActiveDirectory: `Disable-ADUser`.
- FileSystem: `Get-DiskUsageReport`.
- HyperV: `New-DebianVM`, `New-WindowsVM`.
- M365: `Disable-M365User`, `Get-M365AdminMfaReport`, `Get-M365GroupMember`,
  `Get-M365MailboxSpaceReport`, `Get-M365Room`, `Get-M365RoomList`.
- Network: `Get-MerakiFirmware`.
- Private helpers: `Assert-ExchangeOnlineConnection`, `Assert-MgGraphConnection`,
  `Assert-Module`, `Test-IsAdmin`.

### Changed
- Missing modules: instead of failing (Graph / Exchange Online commands) or
  installing silently (`Disable-M365User`), commands now say which modules are
  missing and ask whether to install them from the PowerShell Gallery.
  `Assert-Module` prompts by default and gains `-Force` to install without
  prompting; `Assert-MgGraphConnection` and `Assert-ExchangeOnlineConnection`
  now use it instead of throwing.
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

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).