# Active Directory School Year Migration

CSV-driven tools for end-of-year student account rollover.

This workflow is designed for schools where students move to new grade OUs and grade/security groups at the end of the year.

## Files

| File | Purpose |
| --- | --- |
| `Test-StudentYearMigrationPlan.ps1` | Read-only validation of a migration CSV before changes. |
| `Invoke-StudentYearMigration.ps1` | Moves student AD accounts to target OUs and updates groups with `-WhatIf` support. |
| `examples/student-migration-plan.csv` | Sanitized input template. |
| `examples/student-migration-audit.csv` | Sanitized audit output. |
| `examples/student-migration-results.csv` | Sanitized migration result output. |

## Requirements

- RSAT Active Directory PowerShell module.
- Permissions to read users, OUs, and groups for the audit script.
- Permissions to move users and update group membership for the migration script.

## CSV Columns

Required:

- `StudentId`
- `SamAccountName`
- `UserPrincipalName`
- `CurrentGrade`
- `NextGrade`
- `TargetOU`
- `AddGroups`
- `RemoveGroups`

Optional:

- `ExpectedCurrentOU`
- `Notes`

Use semicolons inside `AddGroups` and `RemoveGroups` when a student needs multiple group changes.

## Workflow

1. Build the migration CSV from SIS, identity system, or your reviewed working file.
2. Run the audit script and fix every failed row.
3. Run the migration script with `-WhatIf`.
4. Review the result log.
5. Run the migration for the approved batch.
6. Keep real outputs in a private folder.

Audit:

```powershell
.\Test-StudentYearMigrationPlan.ps1 `
  -CsvPath .\examples\student-migration-plan.csv `
  -OutputPath .\examples\student-migration-audit.csv
```

Preview:

```powershell
.\Invoke-StudentYearMigration.ps1 `
  -CsvPath .\examples\student-migration-plan.csv `
  -OutputPath .\examples\student-migration-results.csv `
  -WhatIf
```

Run:

```powershell
.\Invoke-StudentYearMigration.ps1 `
  -CsvPath .\private-output\student-migration-plan.csv `
  -OutputPath .\private-output\student-migration-results.csv
```

## Public Safety Notes

- Do not commit real student rosters.
- Do not publish real OUs, student IDs, or group names.
- Keep audit and result logs private.
- Pilot with a small grade or school before running a full rollover.
