# Security Utilities

Read-only security checks and report helpers.

## Files

| File | Purpose |
| --- | --- |
| `Test-EmailDnsRecords.ps1` | Checks MX, SPF, DMARC, and DKIM selector DNS records for a domain. |
| `examples/email-dns-records.csv` | Sanitized example output. |

## Requirements

- PowerShell with `Resolve-DnsName`.
- Permission to query public DNS.

## Usage

Console output:

```powershell
.\Test-EmailDnsRecords.ps1 -Domain example.com
```

CSV report:

```powershell
.\Test-EmailDnsRecords.ps1 `
  -Domain example.com `
  -DkimSelector selector1 `
  -OutputPath .\examples\email-dns-records.csv
```

## Public Safety Notes

- Run only against domains you own or have permission to review.
- Treat customer mail routing and authentication records as sensitive unless approved for release.
