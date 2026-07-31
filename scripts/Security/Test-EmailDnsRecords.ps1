<#
.SYNOPSIS
Check common email DNS records for a domain.

.DESCRIPTION
Checks MX, SPF, DMARC, and optional DKIM selector records using Resolve-DnsName.
Read-only.

.EXAMPLE
.\Test-EmailDnsRecords.ps1 -Domain example.com

.EXAMPLE
.\Test-EmailDnsRecords.ps1 -Domain example.com -DkimSelector selector1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$Domain,

    [string]$DkimSelector = 'selector1'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

function Show-Record {
    param(
        [string]$Label,
        [string]$Name,
        [string]$Type
    )

    Write-Host "`n== $Label =="
    try {
        Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop |
            Select-Object Name, Type, NameExchange, Preference, Strings, IPAddress |
            Format-Table -AutoSize
    } catch {
        Write-Warning "No $Type record found for $Name"
    }
}

Show-Record -Label "MX records" -Name $Domain -Type MX
Show-Record -Label "SPF TXT records" -Name $Domain -Type TXT
Show-Record -Label "DMARC TXT record" -Name "_dmarc.$Domain" -Type TXT
Show-Record -Label "DKIM TXT record" -Name "$DkimSelector._domainkey.$Domain" -Type TXT

Write-Host "`nReview SPF, DMARC, and DKIM records for alignment with your mail providers."
