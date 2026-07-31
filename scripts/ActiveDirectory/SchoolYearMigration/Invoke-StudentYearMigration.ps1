<#
.SYNOPSIS
Move student accounts to new OUs and update grade groups from CSV.

.DESCRIPTION
Processes a reviewed school-year migration CSV. For each student, the script can
move the AD user object to a target OU, add grade/security groups, remove old
grade/security groups, and export a result log. Supports -WhatIf and -Confirm.

.PARAMETER CsvPath
CSV migration plan.

.PARAMETER OutputPath
CSV result log output path.

.PARAMETER Delimiter
CSV delimiter. Defaults to comma.

.PARAMETER Server
Optional domain controller to target.

.PARAMETER Credential
Optional credential for Active Directory cmdlets.

.PARAMETER SkipMove
Do not move users between OUs.

.PARAMETER SkipGroups
Do not update group membership.

.EXAMPLE
.\Invoke-StudentYearMigration.ps1 -CsvPath .\examples\student-migration-plan.csv -OutputPath .\examples\student-migration-results.csv -WhatIf

.EXAMPLE
.\Invoke-StudentYearMigration.ps1 -CsvPath .\examples\student-migration-plan.csv -OutputPath .\private-output\student-migration-results.csv
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [string]$Delimiter = ',',

    [string]$Server,

    [System.Management.Automation.PSCredential]$Credential,

    [switch]$SkipMove,

    [switch]$SkipGroups
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
    $base['Properties'] = @('DistinguishedName','UserPrincipalName','SamAccountName','DisplayName')
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

function Add-Result {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [object]$Row,
        [string]$Action,
        [string]$Target,
        [string]$Status,
        [string]$Message
    )

    $Results.Add([pscustomobject]@{
        Timestamp         = (Get-Date).ToString('s')
        StudentId         = $Row.StudentId
        SamAccountName    = $Row.SamAccountName
        UserPrincipalName = $Row.UserPrincipalName
        CurrentGrade      = $Row.CurrentGrade
        NextGrade         = $Row.NextGrade
        Action            = $Action
        Target            = $Target
        Status            = $Status
        Message           = $Message
    })
}

$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
Assert-CsvColumns -Rows $rows

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    try {
        $user = Get-StudentUser -Row $row
        if (-not $user) {
            Add-Result -Results $results -Row $row -Action 'LookupUser' -Target $row.SamAccountName -Status 'Failed' -Message 'User not found.'
            continue
        }

        if (-not $SkipMove -and -not [string]::IsNullOrWhiteSpace($row.TargetOU)) {
            $moveParams = Get-AdBaseParams
            $moveParams['Identity'] = $user.DistinguishedName
            $moveParams['TargetPath'] = $row.TargetOU
            $moveParams['ErrorAction'] = 'Stop'

            if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Move to $($row.TargetOU)")) {
                Move-ADObject @moveParams
                Add-Result -Results $results -Row $row -Action 'MoveOU' -Target $row.TargetOU -Status 'Success' -Message 'User moved.'
            } else {
                Add-Result -Results $results -Row $row -Action 'MoveOU' -Target $row.TargetOU -Status 'WhatIf' -Message 'Move previewed.'
            }
        }

        if (-not $SkipGroups) {
            foreach ($group in (Split-List -Value $row.RemoveGroups)) {
                $removeParams = Get-AdBaseParams
                $removeParams['Identity'] = $group
                $removeParams['Members'] = $user.DistinguishedName
                $removeParams['Confirm'] = $false
                $removeParams['ErrorAction'] = 'Stop'

                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Remove from group $group")) {
                    Remove-ADGroupMember @removeParams
                    Add-Result -Results $results -Row $row -Action 'RemoveGroup' -Target $group -Status 'Success' -Message 'Removed from group.'
                } else {
                    Add-Result -Results $results -Row $row -Action 'RemoveGroup' -Target $group -Status 'WhatIf' -Message 'Group removal previewed.'
                }
            }

            foreach ($group in (Split-List -Value $row.AddGroups)) {
                $addParams = Get-AdBaseParams
                $addParams['Identity'] = $group
                $addParams['Members'] = $user.DistinguishedName
                $addParams['ErrorAction'] = 'Stop'

                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Add to group $group")) {
                    Add-ADGroupMember @addParams
                    Add-Result -Results $results -Row $row -Action 'AddGroup' -Target $group -Status 'Success' -Message 'Added to group.'
                } else {
                    Add-Result -Results $results -Row $row -Action 'AddGroup' -Target $group -Status 'WhatIf' -Message 'Group add previewed.'
                }
            }
        }
    } catch {
        Add-Result -Results $results -Row $row -Action 'ProcessStudent' -Target $row.SamAccountName -Status 'Failed' -Message $_.Exception.Message
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($results.Count) migration result rows to $OutputPath"
