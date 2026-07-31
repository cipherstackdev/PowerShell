<#
.SYNOPSIS
Export active Microsoft Entra privileged directory role assignments.

.DESCRIPTION
Uses Microsoft Graph PowerShell to export active directory role membership,
including common high-risk roles such as Global Administrator, Privileged Role
Administrator, Exchange Administrator, Security Administrator, and User
Administrator. This report is intended for access reviews and security cleanup.

This script reports active directory role membership. PIM eligible assignment
reporting requires additional role management APIs and may require beta modules
depending on tenant needs.

.PARAMETER OutputPath
CSV output path.

.PARAMETER IncludeAllRoles
Include every active directory role, even roles not in the high-risk list.

.EXAMPLE
.\Export-EntraPrivilegedRoleAssignments.ps1 -OutputPath .\examples\privileged-role-assignments.csv

.EXAMPLE
.\Export-EntraPrivilegedRoleAssignments.ps1 -OutputPath .\examples\all-directory-role-assignments.csv -IncludeAllRoles
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [switch]$IncludeAllRoles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$highRiskRoleNames = @(
    'Global Administrator',
    'Privileged Role Administrator',
    'Exchange Administrator',
    'SharePoint Administrator',
    'Security Administrator',
    'User Administrator',
    'Authentication Administrator',
    'Intune Administrator',
    'Conditional Access Administrator',
    'Application Administrator',
    'Cloud Application Administrator',
    'Helpdesk Administrator'
)

function Assert-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
        throw "Microsoft.Graph.Identity.DirectoryManagement module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'Directory.Read.All','RoleManagement.Read.Directory'"
    }
}

function Get-AdditionalPropertyValue {
    param(
        [object]$InputObject,
        [string]$Name
    )

    if ($null -eq $InputObject) { return '' }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    if ($InputObject.AdditionalProperties -and $InputObject.AdditionalProperties.ContainsKey($Name)) {
        return $InputObject.AdditionalProperties[$Name]
    }
    return ''
}

Assert-GraphConnection

$roles = Get-MgDirectoryRole -All | Sort-Object DisplayName
$rows = New-Object System.Collections.Generic.List[object]

foreach ($role in $roles) {
    $isHighRisk = $role.DisplayName -in $highRiskRoleNames
    if (-not $IncludeAllRoles -and -not $isHighRisk) {
        continue
    }

    $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All)
    if (-not $members) {
        $rows.Add([pscustomobject]@{
            RoleName          = $role.DisplayName
            RoleId            = $role.Id
            RiskTier          = if ($isHighRisk) { 'High' } else { 'Standard' }
            MemberDisplayName = ''
            MemberId          = ''
            MemberType        = ''
            UserPrincipalName = ''
            Mail              = ''
            AccountEnabled    = ''
            Finding           = 'No active members'
        })
        continue
    }

    foreach ($member in $members) {
        $memberType = (Get-AdditionalPropertyValue -InputObject $member -Name '@odata.type') -replace '#microsoft.graph.', ''
        $accountEnabled = Get-AdditionalPropertyValue -InputObject $member -Name 'accountEnabled'
        $upn = Get-AdditionalPropertyValue -InputObject $member -Name 'userPrincipalName'
        $mail = Get-AdditionalPropertyValue -InputObject $member -Name 'mail'

        $finding = if ($isHighRisk) { 'Review privileged assignment' } else { 'Informational' }
        if ($accountEnabled -eq $false) {
            $finding = 'Disabled account has active role'
        }

        $rows.Add([pscustomobject]@{
            RoleName          = $role.DisplayName
            RoleId            = $role.Id
            RiskTier          = if ($isHighRisk) { 'High' } else { 'Standard' }
            MemberDisplayName = Get-AdditionalPropertyValue -InputObject $member -Name 'displayName'
            MemberId          = $member.Id
            MemberType        = $memberType
            UserPrincipalName = $upn
            Mail              = $mail
            AccountEnabled    = $accountEnabled
            Finding           = $finding
        })
    }
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$rows | Sort-Object RiskTier, RoleName, MemberDisplayName | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($rows.Count) privileged role assignment rows to $OutputPath"
