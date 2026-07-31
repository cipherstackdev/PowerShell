<#
.SYNOPSIS
Assign or remove Microsoft 365 licenses from a CSV.

.DESCRIPTION
Reads a CSV containing UserPrincipalName, Action, and SkuPartNumber. Uses
Microsoft Graph PowerShell to assign or remove direct user licenses and writes
a result log. Supports -WhatIf.

.PARAMETER CsvPath
CSV file with UserPrincipalName, Action, and SkuPartNumber columns.

.PARAMETER OutputPath
CSV result log path.

.PARAMETER Delimiter
CSV delimiter. Defaults to comma.

.PARAMETER DefaultUsageLocation
Usage location to set before assignment when a user does not already have one.

.EXAMPLE
.\Update-M365UserLicensesFromCsv.ps1 -CsvPath .\examples\license-changes.csv -OutputPath .\examples\license-change-results.csv -WhatIf
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [string]$Delimiter = ',',

    [ValidatePattern('^[A-Z]{2}$')]
    [string]$DefaultUsageLocation = 'US'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        throw "Microsoft.Graph.Users module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.ReadWrite.All','Organization.Read.All'"
    }
}

function New-ResultRow {
    param(
        [string]$UserPrincipalName,
        [string]$Action,
        [string]$SkuPartNumber,
        [string]$Status,
        [string]$Message
    )

    [pscustomobject]@{
        Timestamp         = (Get-Date).ToString('s')
        UserPrincipalName = $UserPrincipalName
        Action            = $Action
        SkuPartNumber     = $SkuPartNumber
        Status            = $Status
        Message           = $Message
    }
}

function Assert-CsvColumns {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        throw "CSV '$CsvPath' has no data rows."
    }

    $headers = $Rows[0].PSObject.Properties.Name
    $missing = @('UserPrincipalName','Action','SkuPartNumber') | Where-Object { $_ -notin $headers }
    if ($missing) {
        throw "CSV missing required columns: $($missing -join ', '). Present: $($headers -join ', ')"
    }
}

Assert-GraphConnection

$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
Assert-CsvColumns -Rows $rows

$skuMap = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuMap[$_.SkuPartNumber.ToUpperInvariant()] = $_.SkuId
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $upn = ($row.UserPrincipalName).Trim()
    $action = ($row.Action).Trim()
    $skuPartNumber = ($row.SkuPartNumber).Trim().ToUpperInvariant()

    if ($action -notin @('Assign','Remove')) {
        $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'Skipped' -Message 'Action must be Assign or Remove.'))
        continue
    }

    if (-not $skuMap.ContainsKey($skuPartNumber)) {
        $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'Failed' -Message 'SKU was not found in the tenant.'))
        continue
    }

    try {
        $user = Get-MgUser -UserId $upn -Property 'id,userPrincipalName,usageLocation,assignedLicenses' -ErrorAction Stop
        $skuId = $skuMap[$skuPartNumber]

        if ($action -eq 'Assign') {
            if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
                if ($PSCmdlet.ShouldProcess($upn, "Set usage location to $DefaultUsageLocation")) {
                    Update-MgUser -UserId $upn -UsageLocation $DefaultUsageLocation -ErrorAction Stop
                }
            }

            if ($PSCmdlet.ShouldProcess($upn, "Assign license $skuPartNumber")) {
                Set-MgUserLicense -UserId $upn -AddLicenses @(@{ SkuId = $skuId }) -RemoveLicenses @() -ErrorAction Stop | Out-Null
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'Success' -Message 'License assigned.'))
            } else {
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'WhatIf' -Message 'License assignment previewed.'))
            }
        }

        if ($action -eq 'Remove') {
            if ($PSCmdlet.ShouldProcess($upn, "Remove license $skuPartNumber")) {
                Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($skuId) -ErrorAction Stop | Out-Null
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'Success' -Message 'License removed.'))
            } else {
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'WhatIf' -Message 'License removal previewed.'))
            }
        }
    } catch {
        $results.Add((New-ResultRow -UserPrincipalName $upn -Action $action -SkuPartNumber $skuPartNumber -Status 'Failed' -Message $_.Exception.Message))
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($results.Count) license change result rows to $OutputPath"
