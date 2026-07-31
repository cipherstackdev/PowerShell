<#
.SYNOPSIS
Create Microsoft 365 users from a CSV with optional license assignment.

.DESCRIPTION
Reads a CSV of user records, validates required fields, creates users with
Microsoft Graph PowerShell, and optionally assigns a license by SKU part number.
Supports -WhatIf and writes a CSV result log.

.PARAMETER CsvPath
Path to the user import CSV.

.PARAMETER OutputPath
CSV path for the result log.

.PARAMETER Delimiter
CSV delimiter. Defaults to comma.

.PARAMETER DefaultUsageLocation
Fallback usage location when a row does not include UsageLocation.

.PARAMETER AssignLicenses
Assign licenses when LicenseSkuPartNumber is present in the CSV.

.EXAMPLE
.\New-M365UsersFromCsv.ps1 -CsvPath .\examples\users.csv -OutputPath .\examples\import-results.csv -WhatIf

.EXAMPLE
.\New-M365UsersFromCsv.ps1 -CsvPath .\examples\users.csv -OutputPath .\examples\import-results.csv -AssignLicenses
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
    [string]$DefaultUsageLocation = 'US',

    [switch]$AssignLicenses
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredColumns = @(
    'UserPrincipalName',
    'DisplayName',
    'MailNickname',
    'GivenName',
    'Surname',
    'Password',
    'UsageLocation'
)

function Assert-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        throw "Microsoft.Graph.Users module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All','Organization.Read.All'"
    }
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

function Get-LicenseSkuMap {
    $map = @{}
    Get-MgSubscribedSku -All | ForEach-Object {
        $map[$_.SkuPartNumber.ToUpperInvariant()] = $_.SkuId
    }
    return $map
}

function New-ResultRow {
    param(
        [string]$UserPrincipalName,
        [string]$Action,
        [string]$Status,
        [string]$Message
    )

    [pscustomobject]@{
        Timestamp         = (Get-Date).ToString('s')
        UserPrincipalName = $UserPrincipalName
        Action            = $Action
        Status            = $Status
        Message           = $Message
    }
}

Assert-GraphConnection

$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
Assert-CsvColumns -Rows $rows

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$skuMap = if ($AssignLicenses) { Get-LicenseSkuMap } else { @{} }
$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $upn = ($row.UserPrincipalName).Trim()
    $displayName = ($row.DisplayName).Trim()
    $mailNickname = ($row.MailNickname).Trim()
    $usageLocation = if ($row.UsageLocation) { ($row.UsageLocation).Trim().ToUpperInvariant() } else { $DefaultUsageLocation }
    $licenseSku = if ($row.PSObject.Properties.Name -contains 'LicenseSkuPartNumber') { ($row.LicenseSkuPartNumber).Trim().ToUpperInvariant() } else { '' }

    if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($mailNickname)) {
        $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'Validate' -Status 'Skipped' -Message 'Missing required identity values.'))
        continue
    }

    try {
        $existing = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
        if ($existing) {
            $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'CreateUser' -Status 'Skipped' -Message 'User already exists.'))
            continue
        }

        $passwordProfile = @{
            Password = $row.Password
            ForceChangePasswordNextSignIn = $true
        }

        $userParams = @{
            AccountEnabled    = $true
            DisplayName       = $displayName
            MailNickname      = $mailNickname
            UserPrincipalName = $upn
            PasswordProfile   = $passwordProfile
            GivenName         = $row.GivenName
            Surname           = $row.Surname
            UsageLocation     = $usageLocation
        }

        foreach ($optional in 'JobTitle','Department','OfficeLocation','MobilePhone','BusinessPhones') {
            if ($row.PSObject.Properties.Name -contains $optional) {
                $optionalValue = $row.PSObject.Properties[$optional].Value
                if (-not [string]::IsNullOrWhiteSpace($optionalValue)) {
                    if ($optional -eq 'BusinessPhones') {
                        $userParams[$optional] = @($optionalValue)
                    } else {
                        $userParams[$optional] = $optionalValue
                    }
                }
            }
        }

        if ($PSCmdlet.ShouldProcess($upn, 'Create Microsoft 365 user')) {
            [void](New-MgUser @userParams -ErrorAction Stop)
            $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'CreateUser' -Status 'Success' -Message 'User created.'))
        } else {
            $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'CreateUser' -Status 'WhatIf' -Message 'User creation previewed.'))
        }

        if ($AssignLicenses -and $licenseSku) {
            if (-not $skuMap.ContainsKey($licenseSku)) {
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'AssignLicense' -Status 'Failed' -Message "SKU '$licenseSku' was not found in the tenant."))
                continue
            }

            $addLicenses = @(@{ SkuId = $skuMap[$licenseSku] })
            if ($PSCmdlet.ShouldProcess($upn, "Assign license $licenseSku")) {
                Set-MgUserLicense -UserId $upn -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop | Out-Null
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'AssignLicense' -Status 'Success' -Message "Assigned $licenseSku."))
            } else {
                $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'AssignLicense' -Status 'WhatIf' -Message "License assignment previewed for $licenseSku."))
            }
        }
    } catch {
        $results.Add((New-ResultRow -UserPrincipalName $upn -Action 'ProcessUser' -Status 'Failed' -Message $_.Exception.Message))
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($results.Count) result rows to $OutputPath"
