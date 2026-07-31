# Microsoft Entra Group Membership Imports

Add users to Microsoft Entra ID groups from a CSV file.

## Files

| File | Purpose |
| --- | --- |
| `Add-UsersToGroupsFromCsv.ps1` | Adds users to one or more groups from CSV input. |
| `examples/users-to-single-group.csv` | Example where each row targets one group. |
| `examples/users-to-multiple-groups.csv` | Example where one row can target several groups. |

## Requirements

- Microsoft Graph PowerShell module.
- A signed-in Microsoft Graph session with `User.Read.All`, `Group.Read.All`, and `Group.ReadWrite.All`.
- Permission to update the target groups.

Connect first:

```powershell
Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Group.ReadWrite.All"
```

## Examples

Preview a single-group import:

```powershell
.\Add-UsersToGroupsFromCsv.ps1 `
  -CsvPath .\examples\users-to-single-group.csv `
  -WhatIf `
  -Verbose
```

Preview a multi-group import:

```powershell
.\Add-UsersToGroupsFromCsv.ps1 `
  -CsvPath .\examples\users-to-multiple-groups.csv `
  -GroupSeparator ';' `
  -WhatIf `
  -Verbose
```

## CSV Columns

Required:

- `userPrincipalName`
- `group`

The `group` value can be a group display name or group object ID. For multiple groups in one row, separate the group values with `;` unless you pass a different `-GroupSeparator`.

## Public Safety Notes

- Use example.com or placeholder users in documentation.
- Do not commit exported staff lists or production group assignments.
- Test with `-WhatIf` before changing group membership.
