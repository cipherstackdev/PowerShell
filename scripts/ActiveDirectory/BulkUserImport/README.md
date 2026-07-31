# Bulk Active Directory User Import

Create Active Directory users from a CSV with validation, `-WhatIf` support, optional domain controller targeting, and simple result logging.

## Files

| File | Purpose |
| --- | --- |
| `New-BulkAdUsersFromCsv.ps1` | Creates AD users from a CSV. |
| `examples/users.csv` | Sanitized example CSV layout. |

## Requirements

- Windows PowerShell or PowerShell 7 on a machine with RSAT Active Directory tools.
- Active Directory module.
- Permissions to create users in the target OU.

## CSV Columns

Required:

- `givenName`
- `surname`
- `name`
- `displayName`
- `samAccountName`
- `userPrincipalName`
- `password`
- `path`
- `office`
- `title`

Optional:

- `email`

## Examples

Preview changes:

```powershell
.\New-BulkAdUsersFromCsv.ps1 -CsvPath .\examples\users.csv -WhatIf -Verbose
```

Create users and require password change at next logon:

```powershell
.\New-BulkAdUsersFromCsv.ps1 `
  -CsvPath .\examples\users.csv `
  -Server dc01.example.local `
  -ForceChangeAtLogon `
  -LogPath .\created-users.log `
  -Verbose
```

## Public Safety Notes

- The example CSV uses fake names, domains, OUs, and passwords.
- Do not commit real user exports or initial passwords.
- Prefer `-ForceChangeAtLogon` for public examples and production workflows.
- Test with `-WhatIf` before creating users.
