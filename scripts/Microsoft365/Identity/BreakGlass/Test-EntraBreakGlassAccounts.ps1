<#
.SYNOPSIS
Review expected Microsoft Entra break-glass accounts.

.DESCRIPTION
Reads a CSV of expected emergency access accounts and exports a review report
with account state, assigned licenses, active directory roles, and basic safety
findings. This is read-only and intended for periodic emergency access reviews.

.PARAMETER CsvPath
CSV path containing UserPrincipalName and Owner columns.

.PARAMETER OutputPath
CSV output path.

.EXAMPLE
.\Test-EntraBreakGlassAccounts.ps1 -CsvPath .\examples\break-glass-accounts.csv -OutputPath .\examples\break-glass-review.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$OutputPath
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
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All','RoleManagement.Read.Directory','Organization.Read.All'"
    }
}

function Assert-CsvColumns {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -eq 0) {
        throw "CSV '$CsvPath' has no data rows."
    }

    $headers = $Rows[0].PSObject.Properties.Name
    $missing = @('UserPrincipalName','Owner') | Where-Object { $_ -notin $headers }
    if ($missing) {
        throw "CSV missing required columns: $($missing -join ', '). Present: $($headers -join ', ')"
    }
}

function Join-Value {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return ($Value -join ';') }
    return [string]$Value
}

function Get-RoleMembershipMap {
    $map = @{}
    $roles = Get-MgDirectoryRole -All
    foreach ($role in $roles) {
        $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All)
        foreach ($member in $members) {
            if (-not $map.ContainsKey($member.Id)) {
                $map[$member.Id] = New-Object System.Collections.Generic.List[string]
            }
            $map[$member.Id].Add($role.DisplayName)
        }
    }
    return $map
}

Assert-GraphConnection

$rows = Import-Csv -Path $CsvPath
Assert-CsvColumns -Rows $rows

$skuLookup = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuLookup[$_.SkuId.ToString()] = $_.SkuPartNumber
}

$roleMap = Get-RoleMembershipMap
$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $upn = ($row.UserPrincipalName).Trim()
    $owner = ($row.Owner).Trim()
    $notes = if ($row.PSObject.Properties.Name -contains 'Notes') { $row.Notes } else { '' }

    try {
        $user = Get-MgUser -UserId $upn -Property 'id,displayName,userPrincipalName,accountEnabled,createdDateTime,signInActivity,assignedLicenses,userType' -ErrorAction Stop
        $licenseNames = foreach ($license in @($user.AssignedLicenses)) {
            $skuId = $license.SkuId.ToString()
            if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { $skuId }
        }
        $roles = if ($roleMap.ContainsKey($user.Id)) { @($roleMap[$user.Id]) } else { @() }
        $lastSignIn = if ($user.SignInActivity) { $user.SignInActivity.LastSignInDateTime } else { '' }

        $findings = New-Object System.Collections.Generic.List[string]
        if (-not $user.AccountEnabled) { $findings.Add('Account disabled') }
        if (-not $roles) { $findings.Add('No active directory roles found') }
        if ($licenseNames) { $findings.Add('License assigned; review whether emergency account should be unlicensed') }
        if ($lastSignIn) { $findings.Add('Recent sign-in value present; confirm expected emergency use') }
        if ([string]::IsNullOrWhiteSpace($owner)) { $findings.Add('Owner missing') }
        if ($findings.Count -eq 0) { $findings.Add('Review') }

        $results.Add([pscustomobject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName       = $user.DisplayName
            Owner             = $owner
            AccountEnabled    = $user.AccountEnabled
            UserType          = $user.UserType
            CreatedDateTime   = $user.CreatedDateTime
            LastSignIn        = $lastSignIn
            AssignedLicenses  = Join-Value -Value $licenseNames
            ActiveRoles       = Join-Value -Value $roles
            Finding           = Join-Value -Value $findings
            Notes             = $notes
        })
    } catch {
        $results.Add([pscustomobject]@{
            UserPrincipalName = $upn
            DisplayName       = ''
            Owner             = $owner
            AccountEnabled    = ''
            UserType          = ''
            CreatedDateTime   = ''
            LastSignIn        = ''
            AssignedLicenses  = ''
            ActiveRoles       = ''
            Finding           = "Account lookup failed: $($_.Exception.Message)"
            Notes             = $notes
        })
    }
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$results | Sort-Object UserPrincipalName | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($results.Count) break-glass review rows to $OutputPath"
