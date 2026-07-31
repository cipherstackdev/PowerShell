# Mailbox Rule Change Audit Reports

Export Microsoft Purview unified audit log records for mailbox inbox rule changes.

This is useful when investigating requests such as:

- "Who created a mailbox forwarding rule?"
- "When was an inbox rule changed?"
- "Did a suspicious rule delete, redirect, or move mail?"
- "What rule changes happened during this date range?"

## Files

| File | Purpose |
| --- | --- |
| `Export-MailboxRuleChangeAudit.ps1` | Searches the unified audit log and exports normalized mailbox rule change records to CSV. |
| `examples/mailbox-rule-changes.csv` | Sanitized example output. |

## Requirements

- Exchange Online PowerShell module.
- Permission to search the unified audit log, such as membership in a role group that includes Audit Logs or View-Only Audit Logs.
- Audit logging must be enabled and within your tenant's retention window.

Install the module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Examples

Last 30 days:

```powershell
.\Export-MailboxRuleChangeAudit.ps1 `
  -StartDate (Get-Date).AddDays(-30) `
  -EndDate (Get-Date) `
  -OutputPath .\examples\mailbox-rule-changes.csv
```

Specific mailbox and date range:

```powershell
.\Export-MailboxRuleChangeAudit.ps1 `
  -StartDate '2026-07-01' `
  -EndDate '2026-07-31 23:59:59' `
  -UserIds user@example.com `
  -OutputPath .\examples\user-rule-changes.csv
```

## Operations Searched

- `New-InboxRule`
- `Set-InboxRule`
- `Remove-InboxRule`
- `Enable-InboxRule`
- `Disable-InboxRule`
- `UpdateInboxRules`

## Suggested Triage

- Look for rule creation shortly after suspicious sign-in activity.
- Review rules that redirect, forward, move, or delete messages.
- Confirm whether rule changes were made by the mailbox owner, delegate, admin, or compromised session.
- Compare this report with the current-state mailbox forwarding and inbox-rule report under `scripts/Exchange/MailboxRules`.

## Public Safety Notes

- Do not commit exported audit reports.
- Audit results can include usernames, IP addresses, mailbox addresses, and rule details.
- Use sanitized screenshots or sample data when documenting findings publicly.
- Confirm date ranges and time zones with the requestor before treating a report as complete.
