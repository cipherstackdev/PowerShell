# Microsoft 365 Security Reports

Security posture exports for Microsoft 365 and Microsoft Entra ID.

## Files

| File | Purpose |
| --- | --- |
| `Export-MfaConditionalAccessReadiness.ps1` | Exports MFA registration readiness and Conditional Access policy summaries. |
| `examples/mfa-readiness.csv` | Sanitized MFA readiness output. |
| `examples/conditional-access-policies.csv` | Sanitized Conditional Access policy output. |

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph connection with `Reports.Read.All`, `Policy.Read.All`, and `User.Read.All`.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes 'Reports.Read.All','Policy.Read.All','User.Read.All'
```

## Usage

```powershell
.\Export-MfaConditionalAccessReadiness.ps1 `
  -UserOutputPath .\examples\mfa-readiness.csv `
  -PolicyOutputPath .\examples\conditional-access-policies.csv
```

Include disabled users:

```powershell
.\Export-MfaConditionalAccessReadiness.ps1 `
  -UserOutputPath .\examples\mfa-readiness.csv `
  -PolicyOutputPath .\examples\conditional-access-policies.csv `
  -IncludeDisabledUsers
```

## Public Safety Notes

- Do not commit production readiness exports.
- Reports reveal users, admin status, authentication methods, policy names, and group/application targeting.
- Use sanitized samples when documenting security work publicly.
