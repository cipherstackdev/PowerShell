# Microsoft 365 Licensing Reports

License reporting tools for cleanup, renewal, and assignment review.

## Files

| File | Purpose |
| --- | --- |
| `Get-M365LicenseAssignmentReport.ps1` | Exports user license assignments with SKU names and direct/group assignment source. |
| `Update-M365UserLicensesFromCsv.ps1` | Assigns or removes direct user licenses from CSV with `-WhatIf` support. |
| `examples/license-assignments.csv` | Sanitized example report. |
| `examples/license-changes.csv` | Sanitized assignment/removal input template. |
| `examples/license-change-results.csv` | Sanitized result log. |

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph connection with `User.Read.All` and `Organization.Read.All`.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes 'User.Read.All','Organization.Read.All'
```

## Usage

Enabled, licensed users only:

```powershell
.\Get-M365LicenseAssignmentReport.ps1 `
  -OutputPath .\examples\license-assignments.csv
```

Include disabled and unlicensed users:

```powershell
.\Get-M365LicenseAssignmentReport.ps1 `
  -OutputPath .\examples\license-assignments.csv `
  -IncludeDisabledUsers `
  -IncludeUnlicensedUsers
```

Preview license changes:

```powershell
.\Update-M365UserLicensesFromCsv.ps1 `
  -CsvPath .\examples\license-changes.csv `
  -OutputPath .\examples\license-change-results.csv `
  -WhatIf
```

## Review Ideas

- Find disabled users that still have direct license assignments.
- Find users with license assignment errors.
- Separate direct assignments from group-based licensing.
- Confirm usage location is populated before license assignment.
- Use `Update-M365UserLicensesFromCsv.ps1` for direct user license changes only; prefer group-based licensing when that is the tenant standard.

## Public Safety Notes

- Do not commit tenant exports.
- Reports include user identity, department, job title, and license state.
