# Script Index

Quick map of the tools in this repo.

## Microsoft 365 And Entra ID

| Area | Script | Use When |
| --- | --- | --- |
| Cloud user onboarding | `scripts/Microsoft365/Identity/BulkUserImport/New-M365UsersFromCsv.ps1` | Creating Microsoft 365 users from a reviewed CSV import. |
| Privileged roles | `scripts/Microsoft365/Identity/PrivilegedRoles/Export-EntraPrivilegedRoleAssignments.ps1` | Reviewing active Entra privileged role membership. |
| Group membership | `scripts/Microsoft365/GroupMembership/Get-M365GroupMembershipReport.ps1` | Exporting Microsoft 365 group membership for review. |
| Licensing | `scripts/Microsoft365/Licensing/Get-M365LicenseAssignmentReport.ps1` | Finding license assignments, source, and errors. |
| Licensing | `scripts/Microsoft365/Licensing/Update-M365UserLicensesFromCsv.ps1` | Assigning or removing direct user licenses from a reviewed CSV. |
| MFA and CA | `scripts/Microsoft365/Security/Export-MfaConditionalAccessReadiness.ps1` | Checking MFA registration readiness and Conditional Access policy coverage. |

## Exchange Online And Purview

| Area | Script | Use When |
| --- | --- | --- |
| Mailbox holds | `scripts/Exchange/LitigationHoldAndAuditing/Set-MailboxHoldAndAudit.ps1` | Applying hold, audit, retention, and archive settings to a mailbox. |
| Mailbox rules | `scripts/Exchange/MailboxRules/Export-MailboxForwardingAndInboxRules.ps1` | Reviewing current mailbox forwarding and suspicious inbox rules. |
| Audit history | `scripts/Purview/MailboxRuleChanges/Export-MailboxRuleChangeAudit.ps1` | Investigating when inbox rules were created, changed, enabled, disabled, or removed. |

## Active Directory And Azure/Entra Groups

| Area | Script | Use When |
| --- | --- | --- |
| Active Directory | `scripts/ActiveDirectory/BulkUserImport/New-BulkAdUsersFromCsv.ps1` | Creating on-prem AD users from CSV. |
| Entra groups | `scripts/Azure/GroupMembership/Add-UsersToGroupsFromCsv.ps1` | Adding users to Entra ID groups from CSV. |

## Security

| Area | Script | Use When |
| --- | --- | --- |
| Email DNS | `scripts/Security/Test-EmailDnsRecords.ps1` | Checking MX, SPF, DMARC, and DKIM selector records. |

## Safety Pattern

- Read-only scripts should export CSV evidence without changing tenant state.
- Change scripts should support `-WhatIf` wherever PowerShell supports it.
- Examples should be sanitized and use `example.com` style data.
- Real tenant exports belong in `private-output/`, `exports/`, `reports/`, or another ignored/private folder.
