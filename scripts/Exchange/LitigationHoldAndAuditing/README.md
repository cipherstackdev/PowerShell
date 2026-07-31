# Exchange Litigation Hold And Mailbox Auditing

Apply litigation hold, mailbox auditing, single item recovery, deleted item retention, and optional archive enablement for a mailbox.

## Files

| File | Purpose |
| --- | --- |
| `Set-MailboxHoldAndAudit.ps1` | Applies hold/audit settings and prints a verification snapshot. |
| `examples/verification-snapshot.txt` | Sanitized example verification output. |

## Requirements

- Exchange Online PowerShell module.
- Exchange administrative permissions.
- Mailbox licensing that supports Litigation Hold, such as Exchange Online Plan 2 or qualifying Microsoft 365 suites.

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
```

## Usage

Indefinite hold:

```powershell
.\Set-MailboxHoldAndAudit.ps1 `
  -UserPrincipalName user@example.com `
  -CaseNumber CASE-2026-001 `
  -HoldOwner "Legal Team" `
  -HoldComment "Legal hold initiated"
```

Timed hold without archive changes:

```powershell
.\Set-MailboxHoldAndAudit.ps1 `
  -UserPrincipalName user@example.com `
  -CaseNumber CASE-2026-002 `
  -HoldDurationDays 365 `
  -SkipArchive
```

## Public Safety Notes

- Do not commit transcripts or real verification output.
- Legal hold actions should be approved by the proper business/legal authority.
- Confirm licensing and retention requirements before applying changes.
