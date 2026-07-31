# Microsoft Entra Break-Glass Account Review

Review expected emergency access accounts for account state, active role membership, license assignment, and basic review findings.

## Files

| File | Purpose |
| --- | --- |
| `Test-EntraBreakGlassAccounts.ps1` | Exports a read-only review of expected break-glass accounts. |
| `examples/break-glass-accounts.csv` | Sanitized input template. |
| `examples/break-glass-review.csv` | Sanitized example output. |

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph connection with `User.Read.All`, `Directory.Read.All`, `RoleManagement.Read.Directory`, and `Organization.Read.All`.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All','RoleManagement.Read.Directory','Organization.Read.All'
```

## Usage

```powershell
.\Test-EntraBreakGlassAccounts.ps1 `
  -CsvPath .\examples\break-glass-accounts.csv `
  -OutputPath .\examples\break-glass-review.csv
```

## Review Ideas

- Confirm each emergency account is enabled, monitored, and owned.
- Confirm active privileged roles match the emergency access design.
- Review whether licenses are assigned intentionally.
- Confirm recent sign-in activity was expected.
- Store real review output privately.

## Public Safety Notes

- Do not commit real emergency account names.
- Break-glass account details reveal sensitive tenant recovery design.
- Keep review evidence in a private evidence folder.
