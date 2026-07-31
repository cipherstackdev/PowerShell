# Microsoft 365 Bulk User Import

Create Microsoft 365 users from CSV with Microsoft Graph PowerShell.

## Files

| File | Purpose |
| --- | --- |
| `New-M365UsersFromCsv.ps1` | Creates cloud users and optionally assigns licenses. |
| `examples/users.csv` | Sanitized import template. |
| `examples/import-results.csv` | Sanitized sample result log. |

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph connection with delegated scopes such as `User.ReadWrite.All`, `Directory.ReadWrite.All`, and `Organization.Read.All`.
- A test batch before any production run.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All','Organization.Read.All'
```

## Usage

Preview:

```powershell
.\New-M365UsersFromCsv.ps1 `
  -CsvPath .\examples\users.csv `
  -OutputPath .\examples\import-results.csv `
  -WhatIf
```

Create users and assign licenses from `LicenseSkuPartNumber`:

```powershell
.\New-M365UsersFromCsv.ps1 `
  -CsvPath .\examples\users.csv `
  -OutputPath .\examples\import-results.csv `
  -AssignLicenses
```

## CSV Columns

Required:

- `UserPrincipalName`
- `DisplayName`
- `MailNickname`
- `GivenName`
- `Surname`
- `Password`
- `UsageLocation`

Optional:

- `LicenseSkuPartNumber`
- `JobTitle`
- `Department`
- `OfficeLocation`
- `MobilePhone`
- `BusinessPhones`

## Public Safety Notes

- Do not commit real user imports or generated passwords.
- Treat import logs as sensitive.
- Prefer `-WhatIf` and a pilot group before production changes.
