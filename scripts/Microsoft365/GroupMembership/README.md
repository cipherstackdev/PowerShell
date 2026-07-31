# Microsoft 365 Group Membership Reports

Export Microsoft 365 group membership details to CSV.

## Files

| File | Purpose |
| --- | --- |
| `Get-M365GroupMembershipReport.ps1` | Exports group membership to CSV. |
| `examples/group-members.csv` | Sanitized example output. |

## Requirements

- Microsoft Graph PowerShell module.
- A signed-in Microsoft Graph session with `Group.Read.All` and `User.Read.All`.

Connect first:

```powershell
Connect-MgGraph -Scopes "Group.Read.All","User.Read.All"
```

## Example

```powershell
.\Get-M365GroupMembershipReport.ps1 `
  -Group "All Staff" `
  -OutputPath .\examples\group-members.csv
```

## Public Safety Notes

- Do not commit real group exports.
- Group reports can include names, mailbox addresses, and object IDs.
- Share sanitized examples publicly and keep real reports in private case notes.
