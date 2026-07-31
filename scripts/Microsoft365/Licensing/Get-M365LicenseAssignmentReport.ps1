<#
.SYNOPSIS
Export Microsoft 365 license assignment details for users.

.DESCRIPTION
Uses Microsoft Graph PowerShell to export assigned licenses and license assignment
states. The report is useful before license cleanup, true-up work, or migration
planning.

.PARAMETER OutputPath
CSV output path.

.PARAMETER IncludeDisabledUsers
Include disabled accounts in the report.

.PARAMETER IncludeUnlicensedUsers
Include users with no assigned licenses.

.EXAMPLE
.\Get-M365LicenseAssignmentReport.ps1 -OutputPath .\examples\license-assignments.csv

.EXAMPLE
.\Get-M365LicenseAssignmentReport.ps1 -OutputPath .\examples\license-assignments.csv -IncludeDisabledUsers -IncludeUnlicensedUsers
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [switch]$IncludeDisabledUsers,

    [switch]$IncludeUnlicensedUsers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        throw "Microsoft.Graph.Users module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.Read.All','Organization.Read.All'"
    }
}

function Join-Value {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return ($Value -join ';') }
    return [string]$Value
}

Assert-GraphConnection

$skuLookup = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuLookup[$_.SkuId.ToString()] = $_.SkuPartNumber
}

$properties = @(
    'id',
    'displayName',
    'userPrincipalName',
    'accountEnabled',
    'department',
    'jobTitle',
    'usageLocation',
    'assignedLicenses',
    'licenseAssignmentStates'
)

$users = Get-MgUser -All -Property $properties
$rows = New-Object System.Collections.Generic.List[object]

foreach ($user in $users) {
    if (-not $IncludeDisabledUsers -and -not $user.AccountEnabled) {
        continue
    }

    $assignedLicenses = @($user.AssignedLicenses)
    $assignmentStates = @($user.LicenseAssignmentStates)

    if (-not $assignedLicenses -and -not $IncludeUnlicensedUsers) {
        continue
    }

    if (-not $assignedLicenses) {
        $rows.Add([pscustomobject]@{
            DisplayName             = $user.DisplayName
            UserPrincipalName       = $user.UserPrincipalName
            AccountEnabled          = $user.AccountEnabled
            Department              = $user.Department
            JobTitle                = $user.JobTitle
            UsageLocation           = $user.UsageLocation
            SkuPartNumber           = ''
            SkuId                   = ''
            AssignmentSource        = 'Unlicensed'
            AssignedByGroup         = ''
            DisabledPlans           = ''
            Error                   = ''
            LastUpdatedDateTime     = ''
        })
        continue
    }

    foreach ($license in $assignedLicenses) {
        $skuId = $license.SkuId.ToString()
        $matchingStates = @($assignmentStates | Where-Object { $_.SkuId.ToString() -eq $skuId })
        if (-not $matchingStates) {
            $matchingStates = @($null)
        }

        foreach ($state in $matchingStates) {
            $assignedByGroup = if ($state) { $state.AssignedByGroup } else { '' }
            $source = if ($assignedByGroup) { 'Group' } else { 'Direct' }
            $rows.Add([pscustomobject]@{
                DisplayName             = $user.DisplayName
                UserPrincipalName       = $user.UserPrincipalName
                AccountEnabled          = $user.AccountEnabled
                Department              = $user.Department
                JobTitle                = $user.JobTitle
                UsageLocation           = $user.UsageLocation
                SkuPartNumber           = if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { $skuId }
                SkuId                   = $skuId
                AssignmentSource        = $source
                AssignedByGroup         = $assignedByGroup
                DisabledPlans           = Join-Value -Value $license.DisabledPlans
                Error                   = if ($state) { $state.Error } else { '' }
                LastUpdatedDateTime     = if ($state) { $state.LastUpdatedDateTime } else { '' }
            })
        }
    }
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$rows | Sort-Object UserPrincipalName, SkuPartNumber, AssignmentSource | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($rows.Count) license assignment rows to $OutputPath"
