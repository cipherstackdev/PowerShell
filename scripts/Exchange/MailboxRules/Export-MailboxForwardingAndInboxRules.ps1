<#
.SYNOPSIS
Export mailbox forwarding settings and inbox rules for Exchange Online mailboxes.

.DESCRIPTION
Connects to Exchange Online PowerShell, exports mailbox forwarding configuration,
and inventories inbox rules with forwarding, redirect, delete, and move actions.
Useful for compromise review, helpdesk requests, and periodic mailbox hygiene.

.PARAMETER Mailboxes
Optional mailbox identities to inspect. If omitted, all user mailboxes are reviewed.

.PARAMETER OutputDirectory
Folder for CSV outputs.

.PARAMETER IncludeAllRules
Include non-suspicious inbox rules in the inbox rule export.

.EXAMPLE
.\Export-MailboxForwardingAndInboxRules.ps1 -OutputDirectory .\examples

.EXAMPLE
.\Export-MailboxForwardingAndInboxRules.ps1 -Mailboxes user@example.com -OutputDirectory .\examples -IncludeAllRules
#>

[CmdletBinding()]
param(
    [string[]]$Mailboxes,

    [Parameter(Mandatory)]
    [string]$OutputDirectory,

    [switch]$IncludeAllRules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Connect-ExchangeOnlineIfNeeded {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connected = $false
    try {
        $connection = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        if ($connection) { $connected = $true }
    } catch {
        $connected = $false
    }

    if (-not $connected) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

function Join-ExchangeValue {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return ($Value -join ';') }
    return [string]$Value
}

function Test-SuspiciousRule {
    param([object]$Rule)

    return [bool](
        $Rule.ForwardTo -or
        $Rule.ForwardAsAttachmentTo -or
        $Rule.RedirectTo -or
        $Rule.DeleteMessage -or
        $Rule.SoftDeleteMessage -or
        $Rule.MoveToFolder
    )
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

Connect-ExchangeOnlineIfNeeded

$mailboxList = if ($Mailboxes) {
    foreach ($mailbox in $Mailboxes) {
        Get-EXOMailbox -Identity $mailbox -Properties ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward -ErrorAction Stop
    }
} else {
    Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward
}

$forwardingRows = New-Object System.Collections.Generic.List[object]
$ruleRows = New-Object System.Collections.Generic.List[object]

foreach ($mailbox in $mailboxList) {
    $mailboxIdentity = $mailbox.UserPrincipalName
    if (-not $mailboxIdentity) { $mailboxIdentity = $mailbox.PrimarySmtpAddress }

    $forwardingRows.Add([pscustomobject]@{
        DisplayName                = $mailbox.DisplayName
        UserPrincipalName          = $mailboxIdentity
        PrimarySmtpAddress         = $mailbox.PrimarySmtpAddress
        RecipientTypeDetails       = $mailbox.RecipientTypeDetails
        ForwardingSmtpAddress      = $mailbox.ForwardingSmtpAddress
        ForwardingAddress          = $mailbox.ForwardingAddress
        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        ForwardingFinding          = if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { 'Forwarding configured' } else { 'None' }
    })

    try {
        $rules = Get-InboxRule -Mailbox $mailboxIdentity -ErrorAction Stop
    } catch {
        $ruleRows.Add([pscustomobject]@{
            Mailbox               = $mailboxIdentity
            RuleName              = ''
            Enabled               = ''
            Priority              = ''
            ForwardTo             = ''
            ForwardAsAttachmentTo = ''
            RedirectTo            = ''
            DeleteMessage         = ''
            SoftDeleteMessage     = ''
            MoveToFolder          = ''
            StopProcessingRules   = ''
            Finding               = 'Rule read failed'
            Error                 = $_.Exception.Message
        })
        continue
    }

    foreach ($rule in $rules) {
        $isSuspicious = Test-SuspiciousRule -Rule $rule
        if (-not $IncludeAllRules -and -not $isSuspicious) {
            continue
        }

        $ruleRows.Add([pscustomobject]@{
            Mailbox               = $mailboxIdentity
            RuleName              = $rule.Name
            Enabled               = $rule.Enabled
            Priority              = $rule.Priority
            ForwardTo             = Join-ExchangeValue -Value $rule.ForwardTo
            ForwardAsAttachmentTo = Join-ExchangeValue -Value $rule.ForwardAsAttachmentTo
            RedirectTo            = Join-ExchangeValue -Value $rule.RedirectTo
            DeleteMessage         = $rule.DeleteMessage
            SoftDeleteMessage     = $rule.SoftDeleteMessage
            MoveToFolder          = $rule.MoveToFolder
            StopProcessingRules   = $rule.StopProcessingRules
            Finding               = if ($isSuspicious) { 'Review rule action' } else { 'Informational' }
            Error                 = ''
        })
    }
}

$forwardingPath = Join-Path -Path $OutputDirectory -ChildPath 'mailbox-forwarding.csv'
$rulesPath = Join-Path -Path $OutputDirectory -ChildPath 'inbox-rules.csv'

$forwardingRows | Sort-Object UserPrincipalName | Export-Csv -Path $forwardingPath -NoTypeInformation -Encoding UTF8
$ruleRows | Sort-Object Mailbox, Priority | Export-Csv -Path $rulesPath -NoTypeInformation -Encoding UTF8

Write-Host "Wrote forwarding report to $forwardingPath"
Write-Host "Wrote inbox rule report to $rulesPath"
