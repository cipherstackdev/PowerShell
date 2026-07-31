<#
.SYNOPSIS
Validate a student school-year migration CSV before moving accounts.

.DESCRIPTION
Reads a CSV containing student account moves, target OUs, and group changes.
Validates required columns, confirms users/OUs/groups exist, checks optional
current OU expectations, and exports an audit report. This script is read-only.

.PARAMETER CsvPath
CSV migration plan.

.PARAMETER OutputPath
CSV audit report output path.

.PARAMETER Delimiter
CSV delimiter. Defaults to comma.

.PARAMETER Server
Optional domain controller to target.

.PARAMETER Credential
Optional credential for Active Directory cmdlets.

.EXAMPLE
.\Test-StudentYearMigrationPlan.ps1 -CsvPath .\examples\student-migration-plan.csv -OutputPath .\examples\student-migration-audit.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [string]$Delimiter = ',',

    [string]$Server,

    [System.Management.Automation.PSCredential]$Credential
)

#requires -Modules ActiveDirectory

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredColumns = @(
    'StudentId',
    'SamAccountName',
    'UserPrincipalName',
    'CurrentGrade',
    'NextGrade',
    'TargetOU',
    'AddGroups',
    'RemoveGroups'
)

function Add-IfValue {
    param([hashtable]$Target, [string]$Name, [object]$Value)
    if ($null -ne $Value -and $Value -ne '') { $Target[$Name] = $Value }
}

function Get-AdBaseParams {
    $params = @{}
    Add-IfValue $params 'Server' $Server
    Add-IfValue $params 'Credential' $Credential
    return $params
}

function Assert-CsvColumns {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        throw "CSV '$CsvPath' has no data rows."
    }

    $headers = $Rows[0].PSObject.Properties.Name
    $missing = $requiredColumns | Where-Object { $_ -notin $headers }
    if ($missing) {
        throw "CSV missing required columns: $($missing -join ', '). Present: $($headers -join ', ')"
    }
}

function Split-List {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @($Value -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-StudentUser {
    param([object]$Row)

    $base = Get-AdBaseParams
    $base['Properties'] = @('DistinguishedName','Enabled','MemberOf','UserPrincipalName','SamAccountName','DisplayName')
    $base['ErrorAction'] = 'Stop'

    if (-not [string]::IsNullOrWhiteSpace($Row.SamAccountName)) {
        $base['Identity'] = $Row.SamAccountName
        try { return Get-ADUser @base } catch {}
        $base.Remove('Identity')
    }

    if (-not [string]::IsNullOrWhiteSpace($Row.UserPrincipalName)) {
        $escapedUpn = $Row.UserPrincipalName.Replace("'", "''")
        $base['Filter'] = "UserPrincipalName -eq '$escapedUpn'"
        return Get-ADUser @base
    }

    return $null
}

function Test-AdObjectExists {
    param(
        [string]$Identity,
        [string]$Type
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) { return $false }

    $base = Get-AdBaseParams
    $base['Identity'] = $Identity
    $base['ErrorAction'] = 'Stop'

    try {
        if ($Type -eq 'OU') {
            [void](Get-ADOrganizationalUnit @base)
        } else {
            [void](Get-ADGroup @base)
        }
        return $true
    } catch {
        return $false
    }
}

$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
Assert-CsvColumns -Rows $rows

$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $issues = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $user = Get-StudentUser -Row $row

    if (-not $user) {
        $issues.Add('User not found')
    }

    if (-not (Test-AdObjectExists -Identity $row.TargetOU -Type 'OU')) {
        $issues.Add('Target OU not found')
    }

    if (($row.PSObject.Properties.Name -contains 'ExpectedCurrentOU') -and -not [string]::IsNullOrWhiteSpace($row.ExpectedCurrentOU) -and $user) {
        if ($user.DistinguishedName -notlike "*$($row.ExpectedCurrentOU)") {
            $warnings.Add('Current OU does not match ExpectedCurrentOU')
        }
    }

    foreach ($group in (Split-List -Value $row.AddGroups)) {
        if (-not (Test-AdObjectExists -Identity $group -Type 'Group')) {
            $issues.Add("Add group not found: $group")
        }
    }

    foreach ($group in (Split-List -Value $row.RemoveGroups)) {
        if (-not (Test-AdObjectExists -Identity $group -Type 'Group')) {
            $issues.Add("Remove group not found: $group")
        }
    }

    if ([string]::IsNullOrWhiteSpace($row.NextGrade)) {
        $warnings.Add('NextGrade is blank')
    }

    $results.Add([pscustomobject]@{
        StudentId          = $row.StudentId
        SamAccountName     = $row.SamAccountName
        UserPrincipalName  = $row.UserPrincipalName
        CurrentGrade       = $row.CurrentGrade
        NextGrade          = $row.NextGrade
        TargetOU           = $row.TargetOU
        AddGroups          = $row.AddGroups
        RemoveGroups       = $row.RemoveGroups
        UserFound          = [bool]$user
        CurrentDN          = if ($user) { $user.DistinguishedName } else { '' }
        Enabled            = if ($user) { $user.Enabled } else { '' }
        Status             = if ($issues.Count -gt 0) { 'Fail' } elseif ($warnings.Count -gt 0) { 'Warn' } else { 'Pass' }
        Issues             = ($issues -join '; ')
        Warnings           = ($warnings -join '; ')
    })
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($results.Count) migration audit rows to $OutputPath"
