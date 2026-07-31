<#
.SYNOPSIS
Check common email DNS records for a domain.

.DESCRIPTION
Checks MX, SPF, DMARC, and optional DKIM selector records using Resolve-DnsName.
Read-only. Outputs structured findings and can export a CSV report.

.PARAMETER Domain
Domain to check.

.PARAMETER DkimSelector
DKIM selector to check. Defaults to selector1.

.PARAMETER OutputPath
Optional CSV output path.

.EXAMPLE
.\Test-EmailDnsRecords.ps1 -Domain example.com

.EXAMPLE
.\Test-EmailDnsRecords.ps1 -Domain example.com -DkimSelector selector1 -OutputPath .\examples\email-dns-records.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$Domain,

    [string]$DkimSelector = 'selector1',

    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-RecordSafe {
    param(
        [string]$Name,
        [string]$Type
    )

    try {
        return @(Resolve-DnsName -Name $Name -Type $Type -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Join-DnsStrings {
    param([object[]]$Records)

    $values = foreach ($record in $Records) {
        if ($record.Strings) {
            ($record.Strings -join '')
        } elseif ($record.NameExchange) {
            "$($record.Preference) $($record.NameExchange)"
        } elseif ($record.IPAddress) {
            $record.IPAddress
        }
    }

    return (($values | Where-Object { $_ }) -join '; ')
}

function New-Finding {
    param(
        [string]$RecordType,
        [string]$Name,
        [string]$Status,
        [string]$Values,
        [string]$Finding
    )

    [pscustomobject]@{
        Domain     = $Domain
        RecordType = $RecordType
        Name       = $Name
        Status     = $Status
        Values     = $Values
        Finding    = $Finding
    }
}

$findings = New-Object System.Collections.Generic.List[object]

$mxRecords = Resolve-RecordSafe -Name $Domain -Type MX
$findings.Add((New-Finding -RecordType 'MX' -Name $Domain -Status $(if ($mxRecords) { 'Pass' } else { 'Warn' }) -Values (Join-DnsStrings -Records $mxRecords) -Finding $(if ($mxRecords) { 'MX records found.' } else { 'No MX records found.' })))

$txtRecords = Resolve-RecordSafe -Name $Domain -Type TXT
$spfRecords = @($txtRecords | Where-Object { (Join-DnsStrings -Records @($_)) -match '^v=spf1' })
$spfFinding = if ($spfRecords.Count -eq 0) {
    'No SPF record found.'
} elseif ($spfRecords.Count -gt 1) {
    'Multiple SPF records found; publish only one SPF record.'
} else {
    'SPF record found.'
}
$findings.Add((New-Finding -RecordType 'SPF' -Name $Domain -Status $(if ($spfRecords.Count -eq 1) { 'Pass' } else { 'Warn' }) -Values (Join-DnsStrings -Records $spfRecords) -Finding $spfFinding))

$dmarcName = "_dmarc.$Domain"
$dmarcRecords = Resolve-RecordSafe -Name $dmarcName -Type TXT
$dmarcValues = Join-DnsStrings -Records $dmarcRecords
$dmarcStatus = if ($dmarcValues -match 'v=DMARC1' -and $dmarcValues -match 'p=(quarantine|reject)') { 'Pass' } elseif ($dmarcValues -match 'v=DMARC1') { 'Warn' } else { 'Warn' }
$dmarcFinding = if ($dmarcValues -match 'v=DMARC1' -and $dmarcValues -match 'p=(quarantine|reject)') { 'DMARC enforcement policy found.' } elseif ($dmarcValues -match 'v=DMARC1') { 'DMARC exists but is not enforcing quarantine or reject.' } else { 'No DMARC record found.' }
$findings.Add((New-Finding -RecordType 'DMARC' -Name $dmarcName -Status $dmarcStatus -Values $dmarcValues -Finding $dmarcFinding))

$dkimName = "$DkimSelector._domainkey.$Domain"
$dkimRecords = Resolve-RecordSafe -Name $dkimName -Type TXT
$findings.Add((New-Finding -RecordType 'DKIM' -Name $dkimName -Status $(if ($dkimRecords) { 'Pass' } else { 'Warn' }) -Values (Join-DnsStrings -Records $dkimRecords) -Finding $(if ($dkimRecords) { 'DKIM selector record found.' } else { 'DKIM selector record not found.' })))

if ($OutputPath) {
    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }
    $findings | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote email DNS report to $OutputPath"
} else {
    $findings | Format-Table -AutoSize
}

Write-Host "`nReview SPF, DMARC, and DKIM records for alignment with your mail providers."
