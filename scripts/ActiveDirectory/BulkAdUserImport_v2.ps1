<# 
.SYNOPSIS
Create Active Directory users from a CSV.

.DESCRIPTION
Reads a CSV and creates users in AD. Works as a standalone script (no module needed).
Supports -WhatIf/-Confirm, validates headers, quotes-in-DN safe, and lets you target a DC.

.PARAMETER CsvPath
Path to the CSV file.

.PARAMETER Delimiter
CSV delimiter. Defaults to comma. If you prefer `;`, set -Delimiter ';'.

.PARAMETER Server
Optional AD server/GC/DC to target.

.PARAMETER Credential
Optional credential to run AD cmdlets.

.PARAMETER LogPath
Optional path to append a simple CSV log of results.

.PARAMETER PasswordNeverExpires
If set, marks the password as never expiring.

.PARAMETER ForceChangeAtLogon
If set, requires the user to change password at next logon.

.EXAMPLE
.\New-BulkAdUser.ps1 -CsvPath .\Users.csv -Verbose -WhatIf

.EXAMPLE
.\New-BulkAdUser.ps1 -CsvPath .\Users.csv -Server dc01.domain.local -PasswordNeverExpires -LogPath .\created.log
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [string]$Delimiter = ',',

    [string]$Server,

    [System.Management.Automation.PSCredential]$Credential,

    [string]$LogPath,

    [switch]$PasswordNeverExpires,

    [switch]$ForceChangeAtLogon
)

# Make sure the RSAT AD module is available
# (If this fails, PowerShell will stop before doing anything destructive.)
#requires -Modules ActiveDirectory

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Verbose "Importing CSV from: $CsvPath (Delimiter='$Delimiter')"
try {
    $rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
} catch {
    throw "Failed to import CSV `"$CsvPath`": $($_.Exception.Message)"
}

if (-not $rows -or $rows.Count -eq 0) {
    throw "CSV `"$CsvPath`" has no data rows."
}

# Validate required headers (case-insensitive)
$required = 'givenName','surname','name','displayName','samAccountName','userPrincipalName','password','path','office','title'
$headers  = $rows[0].psobject.Properties.Name
$missing  = $required | Where-Object { $_ -notin $headers }
if ($missing) {
    $hint = "If your DN (path) contains commas, ensure it is wrapped in double quotes or use -Delimiter ';'."
    throw "CSV missing required columns: $($missing -join ', '). Columns seen: $($headers -join ', '). $hint"
}

# Helper to optionally add parameters
function Add-IfValue {
    param([hashtable]$Target,[string]$Name,[object]$Value)
    if ($null -ne $Value -and $Value -ne '') { $Target[$Name] = $Value }
}

# Process each row
foreach ($r in $rows) {
    # Trim commonly used fields; property names are case-insensitive
    $upn    = ($r.userPrincipalName).Trim()
    $sam    = ($r.samAccountName).Trim()
    $name   = ($r.name).Trim()
    $disp   = ($r.displayName).Trim()
    $given  = ($r.givenName).Trim()
    $sn     = ($r.surname).Trim()
    $email  = if ($r.PSObject.Properties.Match('email')) { ($r.email).Trim() } else { $upn }
    $pathDN = ($r.path).Trim()
    $office = $r.office
    $title  = $r.title
    $pwdRaw = $r.password

    if ([string]::IsNullOrWhiteSpace($upn)) { Write-Warning "Skipping row (sam=$sam): missing userPrincipalName"; continue }
    if ([string]::IsNullOrWhiteSpace($sam)) { Write-Warning "Skipping $upn: missing samAccountName"; continue }
    if ([string]::IsNullOrWhiteSpace($pathDN)) { Write-Warning "Skipping $upn: missing path DN"; continue }
    if ([string]::IsNullOrWhiteSpace($pwdRaw)) { Write-Warning "Skipping $upn: missing password"; continue }

    # Validate the OU/DN exists & is reachable
    try {
        $ouParams = @{ Identity = $pathDN; ErrorAction = 'Stop' }
        if ($Server)     { $ouParams.Server     = $Server }
        if ($Credential) { $ouParams.Credential = $Credential }
        [void](Get-ADObject @ouParams)
    } catch {
        Write-Error "Path DN not found or inaccessible for $upn: $pathDN. $($_.Exception.Message)"
        continue
    }

    # Skip if user already exists (check UPN and sAMAccountName)
    $exists = $null
    try {
        $flt = "UserPrincipalName -eq '$upn' -or SamAccountName -eq '$sam'"
        $exists = Get-ADUser -Filter $flt -Server $Server -Credential $Credential -ErrorAction SilentlyContinue
    } catch {
        # Non-fatal; continue to attempt creation
    }
    if ($exists) {
        Write-Warning "Skipping $upn: user already exists ($($exists.DistinguishedName))"
        continue
    }

    # Build parameters for New-ADUser
    $params = @{}
    Add-IfValue $params 'GivenName'           $given
    Add-IfValue $params 'Surname'             $sn
    Add-IfValue $params 'Name'                $name
    Add-IfValue $params 'DisplayName'         $disp
    Add-IfValue $params 'SamAccountName'      $sam
    Add-IfValue $params 'UserPrincipalName'   $upn
    Add-IfValue $params 'EmailAddress'        $email
    Add-IfValue $params 'Path'                $pathDN
    Add-IfValue $params 'Office'              $office
    Add-IfValue $params 'Title'               $title
    Add-IfValue $params 'Server'              $Server
    Add-IfValue $params 'Credential'          $Credential

    $params['AccountPassword'] = ConvertTo-SecureString $pwdRaw -AsPlainText -Force
    $params['Enabled'] = $true
    if ($PasswordNeverExpires) { $params['PasswordNeverExpires'] = $true }
    if ($ForceChangeAtLogon)   { $params['ChangePasswordAtLogon'] = $true }

    if ($PSCmdlet.ShouldProcess($upn, 'Create AD user')) {
        try {
            $user = New-ADUser @params -PassThru -ErrorAction Stop
            Write-Verbose "Created: $($user.DistinguishedName)"

            if ($LogPath) {
                "Created,$upn,$sam,$($user.DistinguishedName)" | Out-File -FilePath $LogPath -Append -Encoding utf8
            }
        } catch {
            Write-Error "Failed to create $upn: $($_.Exception.Message)"
            if ($LogPath) { "Error,$upn,$sam,$($_.Exception.Message)" | Out-File -FilePath $LogPath -Append -Encoding utf8 }
        }
    }
}

Write-Verbose "Done."
