# Changelog

## [0.6.0]

First tracked version. Earlier per-function version numbers have been retired
in favor of a single module version. History from those notes is summarized below.

### Added
- ActiveDirectory: `Disable-ADUser`.
- FileSystem: `Get-DiskUsageReport`.
- HyperV: `New-DebianVM`, `New-WindowsVM`.
- M365: `Disable-M365User`, `Get-M365AdminMfaReport`, `Get-M365GroupMember`,
  `Get-M365LicenseActivity`, `Get-M365MailboxSpaceReport`,
  `Get-M365PasswordLastSet`, `Get-M365Room`, `Get-M365RoomList`,
  `Get-M365UserSummary`.
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

### Fixed
- Microsoft Graph sign-in (`Assert-MgGraphConnection`) no longer fails with a
  `NullReferenceException` in the Windows Web Account Manager (WAM) broker
  (`RuntimeBroker`) in hosts such as Windows Terminal, the VS Code terminal, or
  elevated/remote sessions. WAM is now disabled with
  `Set-MgGraphOption -DisableLoginByWAM $true` before connecting, when the
  installed `Microsoft.Graph.Authentication` supports it. A new `-UseWAM` switch
  leaves the setting unchanged. If sign-in still fails in the broker, a short,
  actionable error is shown instead of the raw MSAL stack trace.
  Note: `Set-MgGraphOption` saves this setting to the Graph SDK profile
  (`~/.mg`), so it also applies to later `Connect-MgGraph` calls outside this
  module. Run `Set-MgGraphOption -DisableLoginByWAM $false` to re-enable WAM.
- Exchange Online sign-in (`Assert-ExchangeOnlineConnection`) now calls
  `Connect-ExchangeOnline` with `-DisableWAM` when the installed
  `ExchangeOnlineManagement` supports it (3.7.2 and later), to avoid the same
  WAM broker `NullReferenceException`. A new `-UseWAM` switch opts back into
  broker sign-in. If sign-in still fails in the broker, a short, actionable
  error is shown instead of the raw MSAL stack trace, including a specific hint
  when Microsoft Graph was already connected in the session.
- `Get-M365UserSummary` now connects to Exchange Online before Microsoft Graph.
  Once `Connect-MgGraph` has signed in, `Connect-ExchangeOnline` can fail in the
  WAM broker even with `-DisableWAM` (seen with `ExchangeOnlineManagement`
  3.10.1). Microsoft.Graph.Authentication is still imported first, so the MSAL
  assembly conflict with Graph does not return.

### Known issues
- `Get-MerakiFirmware` uses the deprecated `/networks/{id}/devices` endpoint.
- Exchange Online sign-in can still fail in the WAM broker if Microsoft Graph
  was already connected earlier in the same PowerShell session (for example by
  running a Graph-only command first). Start a new session, or run
  `Connect-ExchangeOnline -DisableWAM` before any Microsoft Graph command.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
