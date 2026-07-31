# PowerShell

PowerShell tools for Microsoft 365, Active Directory, endpoint administration, and security workflows.

These scripts are public examples. Review each script before use, test in a lab or pilot group, and prefer `-WhatIf` where supported.

## Scripts

| Script | Purpose | Permissions |
| --- | --- | --- |
| `scripts/ActiveDirectory/BulkUserImport/New-BulkAdUsersFromCsv.ps1` | Creates Active Directory users from a CSV with validation and `-WhatIf` support. | AD user creation permissions |
| `scripts/Azure/GroupMembership/Add-UsersToGroupsFromCsv.ps1` | Adds users to one or more Microsoft Entra ID groups from a CSV. | `User.Read.All`, `Group.Read.All`, `Group.ReadWrite.All` |
| `scripts/Microsoft365/GroupMembership/Get-M365GroupMembershipReport.ps1` | Exports Microsoft 365 group membership to CSV using Microsoft Graph. | `Group.Read.All`, `User.Read.All` |
| `scripts/Purview/MailboxRuleChanges/Export-MailboxRuleChangeAudit.ps1` | Exports Purview audit records for mailbox inbox rule changes. | Unified audit log search permissions |
| `scripts/Security/Test-EmailDnsRecords.ps1` | Checks MX, SPF, DMARC, and DKIM DNS records for a domain. | None |

## Examples

- Bulk AD user import examples: `scripts/ActiveDirectory/BulkUserImport/examples/`
- Microsoft Entra group import examples: `scripts/Azure/GroupMembership/examples/`
- Microsoft 365 group report examples: `scripts/Microsoft365/GroupMembership/examples/`
- Mailbox rule change audit examples: `scripts/Purview/MailboxRuleChanges/README.md`

## Public Safety Notes

- Do not commit tenant IDs, production domains, customer names, user exports, access tokens, or transcript output.
- Scripts that change users, groups, devices, or security settings should support `-WhatIf`.
- READMEs should include required modules, scopes, examples, and rollback notes where relevant.
