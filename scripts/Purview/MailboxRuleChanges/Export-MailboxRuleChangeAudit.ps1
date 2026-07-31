<#
.SYNOPSIS
Export Purview audit records for mailbox inbox rule changes.

.DESCRIPTION
Searches the Microsoft Purview unified audit log for mailbox rule change events
and exports normalized records to CSV. Useful for helpdesk, security, and legal
requests asking when mailbox forwarding or inbox rules were created, modified,
enabled, disabled, or removed.

.PARAMETER StartDate
Start of the audit search window.

.PARAMETER EndDate
End of the audit search window.

.PARAMETER UserIds
Optional mailbox/user IDs to filter on.

.PARAMETER OutputPath
CSV output path.

.EXAMPLE
.\Export-MailboxRuleChangeAudit.ps1 -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) -OutputPath .\mailbox-rule-changes.csv

.EXAMPLE
.\Export-MailboxRuleChangeAudit.ps1 -StartDate '2026-07-01' -EndDate '2026-07-31 23:59:59' -UserIds user@example.com -OutputPath .\user-rule-changes.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [datetime]$StartDate,

    [Parameter(Mandatory)]
    [datetime]$EndDate,

    [string[]]$UserIds,

    [Parameter(Mandatory)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$operations = @(
    'New-InboxRule',
    'Set-InboxRule',
    'Remove-InboxRule',
    'Enable-InboxRule',
    'Disable-InboxRule',
    'UpdateInboxRules'
)

$outputColumns = @(
    'CreationDate',
    'Operation',
    'UserId',
    'ResultStatus',
    'ClientIP',
    'Workload',
    'ObjectId',
    'LogonUserDisplayName',
    'MailboxOwnerUPN',
    'RuleName',
    'ForwardTo',
    'ForwardAsAttachmentTo',
    'RedirectTo',
    'DeleteMessage',
    'MoveToFolder',
    'StopProcessingRules',
    'RawParameters'
)

function Connect-ExchangeOnlineIfNeeded {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not found. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connected = $false
    try {
        $connection = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        if ($connection) { $connected = $true }
    } catch {
        $connected = $false
    }

    if (-not $connected) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

function Get-PropertyValue {
    param(
        [object]$InputObject,
        [string]$Name
    )

    if ($null -eq $InputObject) { return $null }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }

    return $null
}

function Get-HashtableValue {
    param(
        [hashtable]$Hashtable,
        [string]$Name
    )

    if ($Hashtable.ContainsKey($Name)) { return $Hashtable[$Name] }

    return $null
}

function Convert-AuditRecord {
    param([object]$Record)

    try {
        $data = $Record.AuditData | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Warning "Skipping audit record with invalid JSON: $($Record.CreationDate) $($Record.Operations)"
        return
    }

    $parameters = @{}

    $rawParameters = Get-PropertyValue -InputObject $data -Name 'Parameters'
    if ($rawParameters) {
        foreach ($parameter in $rawParameters) {
            $name = Get-PropertyValue -InputObject $parameter -Name 'Name'
            if ($name) {
                $parameters[$name] = Get-PropertyValue -InputObject $parameter -Name 'Value'
            }
        }
    }

    [pscustomobject]@{
        CreationDate          = Get-PropertyValue -InputObject $Record -Name 'CreationDate'
        Operation             = Get-PropertyValue -InputObject $Record -Name 'Operations'
        UserId                = Get-PropertyValue -InputObject $Record -Name 'UserIds'
        ResultStatus          = Get-PropertyValue -InputObject $data -Name 'ResultStatus'
        ClientIP              = Get-PropertyValue -InputObject $data -Name 'ClientIP'
        Workload              = Get-PropertyValue -InputObject $data -Name 'Workload'
        ObjectId              = Get-PropertyValue -InputObject $data -Name 'ObjectId'
        LogonUserDisplayName  = Get-PropertyValue -InputObject $data -Name 'LogonUserDisplayName'
        MailboxOwnerUPN       = Get-PropertyValue -InputObject $data -Name 'MailboxOwnerUPN'
        RuleName              = Get-HashtableValue -Hashtable $parameters -Name 'Name'
        ForwardTo             = Get-HashtableValue -Hashtable $parameters -Name 'ForwardTo'
        ForwardAsAttachmentTo = Get-HashtableValue -Hashtable $parameters -Name 'ForwardAsAttachmentTo'
        RedirectTo            = Get-HashtableValue -Hashtable $parameters -Name 'RedirectTo'
        DeleteMessage         = Get-HashtableValue -Hashtable $parameters -Name 'DeleteMessage'
        MoveToFolder          = Get-HashtableValue -Hashtable $parameters -Name 'MoveToFolder'
        StopProcessingRules   = Get-HashtableValue -Hashtable $parameters -Name 'StopProcessingRules'
        RawParameters         = (($parameters.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ')
    }
}

if ($EndDate -lt $StartDate) {
    throw "EndDate must be later than StartDate."
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

Connect-ExchangeOnlineIfNeeded

$sessionId = [guid]::NewGuid().Guid
$allRecords = New-Object System.Collections.Generic.List[object]
$resultSize = 5000

do {
    $params = @{
        StartDate = $StartDate
        EndDate = $EndDate
        Operations = $operations
        SessionId = $sessionId
        SessionCommand = 'ReturnLargeSet'
        ResultSize = $resultSize
    }

    if ($UserIds) {
        $params['UserIds'] = $UserIds
    }

    $records = Search-UnifiedAuditLog @params
    foreach ($record in $records) {
        $allRecords.Add($record)
    }
} while (@($records).Count -eq $resultSize)

$output = $allRecords | ForEach-Object { Convert-AuditRecord -Record $_ } | Sort-Object CreationDate

if (-not $output) {
    Write-Warning "No mailbox rule change audit records found for the selected window."
    $output = @()
}

$outputCount = @($output).Count

if ($outputCount -gt 0) {
    $output | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
} else {
    Set-Content -Path $OutputPath -Value ($outputColumns -join ',') -Encoding UTF8
}

Write-Host "Exported $outputCount mailbox rule change audit records to $OutputPath"
