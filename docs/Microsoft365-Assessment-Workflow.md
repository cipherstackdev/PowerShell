# Microsoft 365 Assessment Workflow

This is a public-safe workflow for using the scripts in this repo during a Microsoft 365 identity and messaging review.

## 1. Connect

Use a dedicated admin workstation and a least-privilege admin account where possible.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

Connect-MgGraph -Scopes `
  'User.Read.All',`
  'User.ReadWrite.All',`
  'Directory.Read.All',`
  'Directory.ReadWrite.All',`
  'RoleManagement.Read.Directory',`
  'Organization.Read.All',`
  'Reports.Read.All',`
  'Policy.Read.All'

Connect-ExchangeOnline
```

## 2. Export Read-Only Evidence

```powershell
.\scripts\Microsoft365\Licensing\Get-M365LicenseAssignmentReport.ps1 `
  -OutputPath .\private-output\license-assignments.csv `
  -IncludeDisabledUsers `
  -IncludeUnlicensedUsers

.\scripts\Microsoft365\Security\Export-MfaConditionalAccessReadiness.ps1 `
  -UserOutputPath .\private-output\mfa-readiness.csv `
  -PolicyOutputPath .\private-output\conditional-access-policies.csv

.\scripts\Microsoft365\Identity\PrivilegedRoles\Export-EntraPrivilegedRoleAssignments.ps1 `
  -OutputPath .\private-output\privileged-role-assignments.csv

.\scripts\Microsoft365\Identity\BreakGlass\Test-EntraBreakGlassAccounts.ps1 `
  -CsvPath .\scripts\Microsoft365\Identity\BreakGlass\examples\break-glass-accounts.csv `
  -OutputPath .\private-output\break-glass-review.csv

.\scripts\Exchange\MailboxRules\Export-MailboxForwardingAndInboxRules.ps1 `
  -OutputDirectory .\private-output\mailbox-rules
```

## 3. Review

- Privileged roles: confirm standing access, disabled-account assignments, and admin group membership.
- Break-glass accounts: confirm owners, active privileged roles, license state, and expected sign-in activity.
- MFA readiness: identify users not registered or not capable before enforcing stricter policies.
- Conditional Access: check disabled/report-only policies, broad exclusions, and legacy authentication blocks.
- Licensing: find disabled users with direct licenses and assignment errors.
- Mailbox rules: review forwarding, redirect, delete, and move rules.

## 4. Change With Preview

Use `-WhatIf` first for user creation and license changes.

```powershell
.\scripts\Microsoft365\Identity\BulkUserImport\New-M365UsersFromCsv.ps1 `
  -CsvPath .\scripts\Microsoft365\Identity\BulkUserImport\examples\users.csv `
  -OutputPath .\private-output\user-import-results.csv `
  -WhatIf

.\scripts\Microsoft365\Licensing\Update-M365UserLicensesFromCsv.ps1 `
  -CsvPath .\scripts\Microsoft365\Licensing\examples\license-changes.csv `
  -OutputPath .\private-output\license-change-results.csv `
  -WhatIf
```

## Public Safety

Keep `private-output` and real CSV imports out of Git. The repo includes sanitized examples only.
