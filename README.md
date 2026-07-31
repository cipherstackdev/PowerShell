# PowerShell

PowerShell tools for Microsoft 365, Active Directory, endpoint administration, and security workflows.

These scripts are public examples. Review each script before use, test in a lab or pilot group, and prefer `-WhatIf` where supported.

## Scripts

| Script | Purpose | Permissions |
| --- | --- | --- |
| `scripts/Microsoft365/Get-M365GroupMembershipReport.ps1` | Exports Microsoft 365 group membership to CSV using Microsoft Graph. | `Group.Read.All`, `User.Read.All` |
| `scripts/Security/Test-EmailDnsRecords.ps1` | Checks MX, SPF, DMARC, and DKIM DNS records for a domain. | None |

## Existing Script Cleanup Recommendations

- Keep `BulkAdUserImport_v2.ps1`; it has better parameters, validation, `-WhatIf`, and logging.
- Remove or archive `BulkAdUserImport_v1.ps1` from the public path. It is too easy to misuse because it creates enabled accounts using plaintext CSV passwords and sets passwords to never expire.
- Rename versioned scripts after they stabilize. Example: `BulkAdUserImport_v2.ps1` -> `New-BulkAdUsersFromCsv.ps1`.
- Move CSV files into an `examples/` folder with placeholder domains only.

## Public Safety Notes

- Do not commit tenant IDs, production domains, customer names, user exports, access tokens, or transcript output.
- Scripts that change users, groups, devices, or security settings should support `-WhatIf`.
- READMEs should include required modules, scopes, examples, and rollback notes where relevant.
