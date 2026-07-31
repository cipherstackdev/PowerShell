# Exchange Mailbox Forwarding And Inbox Rule Audit

Export mailbox forwarding settings and inbox rules that may redirect, delete, or move mail.

## Files

| File | Purpose |
| --- | --- |
| `Export-MailboxForwardingAndInboxRules.ps1` | Exports mailbox forwarding settings and suspicious inbox rule actions. |
| `examples/mailbox-forwarding.csv` | Sanitized mailbox forwarding output. |
| `examples/inbox-rules.csv` | Sanitized inbox rule output. |

## Requirements

- Exchange Online PowerShell module.
- Exchange permissions to read mailbox settings and inbox rules.

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
```

## Usage

All user mailboxes:

```powershell
.\Export-MailboxForwardingAndInboxRules.ps1 `
  -OutputDirectory .\examples
```

Specific mailbox:

```powershell
.\Export-MailboxForwardingAndInboxRules.ps1 `
  -Mailboxes user@example.com `
  -OutputDirectory .\examples `
  -IncludeAllRules
```

## Public Safety Notes

- Do not commit production mailbox exports.
- Reports can include mailbox addresses, forwarding targets, folder names, and rule names.
- Review mailbox-level forwarding and inbox rules together during compromise investigations.
