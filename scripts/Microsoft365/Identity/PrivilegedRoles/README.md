# Microsoft Entra Privileged Role Assignments

Export active Microsoft Entra directory role assignments for privileged access review.

## Files

| File | Purpose |
| --- | --- |
| `Export-EntraPrivilegedRoleAssignments.ps1` | Exports active directory role members with high-risk role tagging. |
| `examples/privileged-role-assignments.csv` | Sanitized example output. |

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph connection with `Directory.Read.All` and `RoleManagement.Read.Directory`.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes 'Directory.Read.All','RoleManagement.Read.Directory'
```

## Usage

High-risk roles:

```powershell
.\Export-EntraPrivilegedRoleAssignments.ps1 `
  -OutputPath .\examples\privileged-role-assignments.csv
```

All active directory roles:

```powershell
.\Export-EntraPrivilegedRoleAssignments.ps1 `
  -OutputPath .\examples\all-directory-role-assignments.csv `
  -IncludeAllRoles
```

## Review Ideas

- Confirm every high-risk role has an owner and business reason.
- Check whether disabled accounts still have active role membership.
- Reduce permanent standing access where possible.
- Compare active role membership against PIM eligible assignments when available.

## Public Safety Notes

- Do not commit tenant exports.
- Role membership reports reveal privileged users, groups, and admin structure.
- This script reports active directory role membership, not a full PIM eligible assignment inventory.
