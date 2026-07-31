<#
.SYNOPSIS
Export Microsoft 365 group membership to CSV.

.DESCRIPTION
Uses Microsoft Graph PowerShell to export group members with basic user details.
Accepts group display names or object IDs. Read-only.

.PARAMETER Group
One or more group display names or object IDs.

.PARAMETER OutputPath
CSV output path.

.EXAMPLE
Connect-MgGraph -Scopes "Group.Read.All","User.Read.All"
.\Get-M365GroupMembershipReport.ps1 -Group "All Staff" -OutputPath .\group-members.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$Group,

    [Parameter(Mandatory)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Groups)) {
    throw "Microsoft.Graph.Groups is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
}

if (-not (Get-MgContext)) {
    throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'Group.Read.All','User.Read.All'"
}

$rows = foreach ($groupInput in $Group) {
    $groupObject = $null

    if ($groupInput -match '^[0-9a-fA-F-]{36}$') {
        $groupObject = Get-MgGroup -GroupId $groupInput -ErrorAction Stop
    } else {
        $escaped = $groupInput.Replace("'", "''")
        $matches = Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop
        if ($matches.Count -gt 1) {
            Write-Warning "Multiple groups named '$groupInput'. Use object ID for exact selection."
        }
        $groupObject = $matches | Select-Object -First 1
    }

    if (-not $groupObject) {
        Write-Warning "Group not found: $groupInput"
        continue
    }

    Get-MgGroupMember -GroupId $groupObject.Id -All | ForEach-Object {
        $member = $_
        [pscustomobject]@{
            GroupDisplayName = $groupObject.DisplayName
            GroupId          = $groupObject.Id
            MemberId         = $member.Id
            MemberType       = $member.AdditionalProperties['@odata.type']
            DisplayName      = $member.AdditionalProperties['displayName']
            UserPrincipalName = $member.AdditionalProperties['userPrincipalName']
            Mail             = $member.AdditionalProperties['mail']
        }
    }
}

$rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($rows.Count) membership rows to $OutputPath"
