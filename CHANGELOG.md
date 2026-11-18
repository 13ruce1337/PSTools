# Changelog

## [0.6.0]

First tracked version. Earlier per-function version numbers have been retired
in favor of a single module version. History from those notes is summarized below.

### Added
- ActiveDirectory: `Disable-ADUser`.
- FileSystem: `Get-DiskUsageReport`.
- HyperV: `New-DebianVM`, `New-WindowsVM`.
- M365: `Disable-M365User`, `Get-M365AdminMfaReport`, `Get-M365GroupMember`,
  `Get-M365License`, `Get-M365MailboxSpaceReport`,
  `Get-M365PasswordLastSet`, `Get-M365Room`, `Get-M365RoomList`,
  `Get-M365User`.
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
  error is shown instead of the raw MSAL stack trace.
- `Get-M365User` connects to Microsoft Graph before Exchange Online, like
  `Disable-M365User`. Once `Connect-ExchangeOnline` has signed in, its MSAL
  assembly is loaded and `Connect-MgGraph` fails with
  `Method not found ... WithLogging` (seen with `Microsoft.Graph.Authentication`
  2.34.0 and `ExchangeOnlineManagement` 3.10.1).
- The MSAL-conflict error from `Assert-MgGraphConnection` now says when Exchange
  Online caused it and no longer suggests a connection order that fails.

### Known issues
- `Get-MerakiFirmware` uses the deprecated `/networks/{id}/devices` endpoint.
- Microsoft Graph sign-in fails with `Method not found` if Exchange Online was
  connected earlier in the same PowerShell session (for example by running
  `Get-M365MailboxSpaceReport` or `Get-M365Room` first). Open a new session, or
  run `Connect-MgGraph` before any Exchange Online command.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
