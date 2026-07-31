<#
.SYNOPSIS
Export MFA registration and Conditional Access readiness data.

.DESCRIPTION
Uses Microsoft Graph PowerShell to export authentication method registration
details and Conditional Access policy summaries. This helps identify users who
are not ready for stricter MFA or phishing-resistant access policies.

.PARAMETER UserOutputPath
CSV path for user registration readiness.

.PARAMETER PolicyOutputPath
CSV path for Conditional Access policy summaries.

.PARAMETER IncludeDisabledUsers
Include disabled accounts in the user readiness report.

.EXAMPLE
.\Export-MfaConditionalAccessReadiness.ps1 -UserOutputPath .\examples\mfa-readiness.csv -PolicyOutputPath .\examples\conditional-access-policies.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$UserOutputPath,

    [Parameter(Mandatory)]
    [string]$PolicyOutputPath,

    [switch]$IncludeDisabledUsers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)) {
        throw "Microsoft.Graph.Reports module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Reports -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    if (-not (Get-MgContext)) {
        throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'Reports.Read.All','Policy.Read.All','User.Read.All'"
    }
}

function Join-Text {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return ($Value -join ';') }
    return [string]$Value
}

function Ensure-ParentDirectory {
    param([string]$Path)
    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}

function Get-NestedValue {
    param(
        [object]$InputObject,
        [string[]]$Path
    )

    $current = $InputObject
    foreach ($part in $Path) {
        if ($null -eq $current) { return $null }
        $property = $current.PSObject.Properties[$part]
        if (-not $property) { return $null }
        $current = $property.Value
    }
    return $current
}

Assert-GraphConnection

$userState = @{}
Get-MgUser -All -Property 'id,userPrincipalName,accountEnabled,department,jobTitle' | ForEach-Object {
    $userState[$_.Id] = $_
}

$registrationRows = Get-MgReportAuthenticationMethodUserRegistrationDetail -All | ForEach-Object {
    $user = if ($userState.ContainsKey($_.Id)) { $userState[$_.Id] } else { $null }
    $accountEnabled = if ($user) { $user.AccountEnabled } else { $null }

    if (-not $IncludeDisabledUsers -and $accountEnabled -eq $false) {
        return
    }

    [pscustomobject]@{
        UserPrincipalName               = $_.UserPrincipalName
        UserDisplayName                 = $_.UserDisplayName
        UserType                        = $_.UserType
        AccountEnabled                  = $accountEnabled
        Department                      = if ($user) { $user.Department } else { '' }
        JobTitle                        = if ($user) { $user.JobTitle } else { '' }
        IsAdmin                         = $_.IsAdmin
        IsMfaRegistered                 = $_.IsMfaRegistered
        IsMfaCapable                    = $_.IsMfaCapable
        IsPasswordlessCapable           = $_.IsPasswordlessCapable
        IsSsprRegistered                = $_.IsSsprRegistered
        IsSsprEnabled                   = $_.IsSsprEnabled
        IsSsprCapable                   = $_.IsSsprCapable
        DefaultMfaMethod                = $_.DefaultMfaMethod
        MethodsRegistered               = Join-Text -Value $_.MethodsRegistered
        SystemPreferredAuthenticationMethods = Join-Text -Value $_.SystemPreferredAuthenticationMethods
        UserPreferredMethodForSecondaryAuthentication = $_.UserPreferredMethodForSecondaryAuthentication
        ReadinessFinding                = if (-not $_.IsMfaRegistered) { 'MFA not registered' } elseif (-not $_.IsMfaCapable) { 'MFA not capable' } elseif (-not $_.IsPasswordlessCapable) { 'Passwordless not capable' } else { 'Ready' }
    }
}

$policyRows = Get-MgIdentityConditionalAccessPolicy -All | ForEach-Object {
    [pscustomobject]@{
        DisplayName          = $_.DisplayName
        State                = $_.State
        CreatedDateTime      = $_.CreatedDateTime
        ModifiedDateTime     = $_.ModifiedDateTime
        IncludedUsers        = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Users','IncludeUsers'))
        ExcludedUsers        = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Users','ExcludeUsers'))
        IncludedGroups       = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Users','IncludeGroups'))
        ExcludedGroups       = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Users','ExcludeGroups'))
        IncludedRoles        = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Users','IncludeRoles'))
        IncludedApplications = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Applications','IncludeApplications'))
        ExcludedApplications = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','Applications','ExcludeApplications'))
        ClientAppTypes       = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('Conditions','ClientAppTypes'))
        GrantControls        = Join-Text -Value (Get-NestedValue -InputObject $_ -Path @('GrantControls','BuiltInControls'))
        Operator             = Get-NestedValue -InputObject $_ -Path @('GrantControls','Operator')
        SessionControls      = if ($_.SessionControls) { ($_.SessionControls | ConvertTo-Json -Compress -Depth 6) } else { '' }
    }
}

Ensure-ParentDirectory -Path $UserOutputPath
Ensure-ParentDirectory -Path $PolicyOutputPath

$registrationRows | Sort-Object ReadinessFinding, UserPrincipalName | Export-Csv -Path $UserOutputPath -NoTypeInformation -Encoding UTF8
$policyRows | Sort-Object State, DisplayName | Export-Csv -Path $PolicyOutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Wrote MFA readiness report to $UserOutputPath"
Write-Host "Wrote Conditional Access policy report to $PolicyOutputPath"
