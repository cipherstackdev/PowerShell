<#
.SYNOPSIS
Bulk-add users to Azure AD groups from a CSV.

.DESCRIPTION
Reads a CSV with userPrincipalName and group columns and ensures each user is a member
of the listed group(s). Groups may be given as DisplayName or ObjectId (GUID).
Supports multiple groups per user using a separator (default ';').

.PARAMETER CsvPath
Path to the CSV file.

.PARAMETER Delimiter
CSV delimiter (default ',').

.PARAMETER GroupSeparator
Separator to split multiple groups in the 'group' column (default ';').

.PARAMETER GroupColumn
Name of the group column (default 'group').

.PARAMETER UPNColumn
Name of the user principal name column (default 'userPrincipalName').

.EXAMPLE
.\Add-UsersToGroupsFromCsv.ps1 -CsvPath .\users-to-groups.csv -Verbose

.EXAMPLE
.\Add-UsersToGroupsFromCsv.ps1 -CsvPath .\users-to-groups.csv -Delimiter ';' -GroupSeparator '|' -Verbose
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [string]$Delimiter = ',',

    [string]$GroupSeparator = ';',

    [string]$GroupColumn = 'group',

    [string]$UPNColumn = 'userPrincipalName'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Requires Graph
# Install-Module Microsoft.Graph -Scope CurrentUser
# Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Group.ReadWrite.All"
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    throw "Microsoft.Graph module not found. Install it: Install-Module Microsoft.Graph -Scope CurrentUser"
}
if (-not (Get-MgContext)) {
    throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.Read.All','Group.Read.All','Group.ReadWrite.All'"
}

# Load CSV
$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
if (-not $rows) { throw "CSV '$CsvPath' has no rows." }

# Validate headers (case-insensitive via property access)
$first = $rows[0]
foreach ($col in @($UPNColumn, $GroupColumn)) {
    if (-not $first.PSObject.Properties.Name -contains $col) {
        throw "CSV missing required column: '$col'. Present: $($first.PSObject.Properties.Name -join ', ')"
    }
}

# Caches to avoid repeated lookups
$groupIdCache   = @{}   # key: input name/id -> value: groupId (GUID)
$memberCache    = @{}   # key: groupId ->