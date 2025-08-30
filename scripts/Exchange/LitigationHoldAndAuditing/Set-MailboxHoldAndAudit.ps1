<# 
Set-MailboxHoldAndAudit.ps1
Place a single Exchange Online mailbox on Litigation Hold, ensure auditing, and apply key retention settings.

Usage:
  pwsh -File .\Set-MailboxHoldAndAudit.ps1 `
    -UserPrincipalName user@domain.com `
    -CaseNumber "CASE-2025-001" `
    -HoldOwner "Legal Team" `
    -HoldComment "Legal hold initiated"

Prereqs:
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  Permissions: Exchange Admin (and appropriate org perms)
  Licensing: EXO Plan 2 (E3/E5/A5, etc.) for Litigation Hold
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidatePattern('^.+@.+\..+$')]
  [string]$UserPrincipalName,

  [string]$CaseNumber = "",
  [string]$HoldOwner = "",
  [string]$HoldComment = "Litigation hold enabled",
  [int]$HoldDurationDays = 0,            # 0/blank = indefinite hold
  [int]$AuditLogDays = 365,              # mailbox audit log retention window
  [int]$DeletedItemsRetentionDays = 30,  # bump from default 14 to 30
  [switch]$SkipArchive                   # skip enabling archive/auto-expanding
)

function Connect-EXOIfNeeded {
  try {
    if (-not (Get-Module ExchangeOnlineManagement)) {
      Import-Module ExchangeOnlineManagement -ErrorAction Stop
    }
    $connected = $false
    try { if (Get-ConnectionInformation -ErrorAction Stop) { $connected = $true } } catch {}
    if (-not $connected) {
      Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
      Connect-ExchangeOnline -ShowProgress:$true
    }
  } catch {
    throw "Could not load/connect ExchangeOnlineManagement. $_"
  }
}

Start-Transcript -Path (Join-Path $PSScriptRoot ("HoldAudit_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")) -ErrorAction SilentlyContinue
Connect-EXOIfNeeded

try {
  $mbx = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
} catch {
  Stop-Transcript | Out-Null
  throw "Mailbox not found or inaccessible: $UserPrincipalName. $_"
}

# 1) Optional capacity prep: archive + auto-expanding archive
if (-not $SkipArchive) {
  try {
    if ($mbx.ArchiveStatus -ne 'Active') {
      Write-Host "Enabling archive mailbox..." -ForegroundColor Cyan
      Enable-Mailbox -Identity $UserPrincipalName -Archive -ErrorAction Stop
    }
    Write-Host "Enabling auto-expanding archive (idempotent)..." -ForegroundColor Cyan
    Enable-Mailbox -Identity $UserPrincipalName -AutoExpandingArchive -ErrorAction SilentlyContinue
  } catch {
    Write-Warning "Archive step: $_"
  }
}

# 2) Single Item Recovery + Deleted Items retention
try {
  if (-not $mbx.SingleItemRecoveryEnabled) {
    Write-Host "Enabling Single Item Recovery..." -ForegroundColor Cyan
    Set-Mailbox -Identity $UserPrincipalName -SingleItemRecoveryEnabled $true -ErrorAction Stop
  }
  if ($mbx.DeletedItemsRetention.Days -lt $DeletedItemsRetentionDays) {
    Write-Host "Setting DeletedItemsRetention to $DeletedItemsRetentionDays days..." -ForegroundColor Cyan
    Set-Mailbox -Identity $UserPrincipalName -DeletedItemsRetention "$DeletedItemsRetentionDays.00:00:00" -ErrorAction Stop
  }
} catch {
  Write-Warning "Retention settings step: $_"
}

# 3) Litigation Hold (+ optional duration/owner/comment with timestamp & case)
try {
  $params = @{
    Identity              = $UserPrincipalName
    LitigationHoldEnabled = $true
  }
  if ($HoldDurationDays -gt 0) { $params['LitigationHoldDuration'] = $HoldDurationDays }
  if ($HoldOwner)              { $params['LitigationHoldOwner']    = $HoldOwner }

  $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss zzz")
  $fullComment = "[${stamp}] $HoldComment"
  if ($CaseNumber) { $fullComment = "$fullComment (Case: $CaseNumber)" }
  $params['LitigationHoldComment'] = $fullComment

  Write-Host "Enabling Litigation Hold..." -ForegroundColor Cyan
  Set-Mailbox @params -ErrorAction Stop
} catch {
  Write-Warning "Litigation Hold step failed: $_"
}

# 4) Mailbox auditing (enforce + extend retention)
try {
  Write-Host "Ensuring mailbox auditing is enabled and extending retention..." -ForegroundColor Cyan
  Set-Mailbox -Identity $UserPrincipalName -AuditEnabled $true -AuditLogAgeLimit "$AuditLogDays.00:00:00" -ErrorAction Stop
} catch {
  Write-Warning "Mailbox auditing step failed: $_"
}

# 5) Nudge Managed Folder Assistant to apply hold promptly
try {
  Write-Host "Starting Managed Folder Assistant..." -ForegroundColor Cyan
  Start-ManagedFolderAssistant -Identity $UserPrincipalName -ErrorAction SilentlyContinue
} catch {
  Write-Warning "Managed Folder Assistant step (non-fatal): $_"
}

# 6) Verification snapshot
try {
  $mbx2  = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
  $stats = Get-MailboxStatistics -Identity $UserPrincipalName -ErrorAction SilentlyContinue
  [PSCustomObject]@{
    UserPrincipalName          = $UserPrincipalName
    LitigationHoldEnabled      = $mbx2.LitigationHoldEnabled
    LitigationHoldDuration     = $mbx2.LitigationHoldDuration
    LitigationHoldOwner        = $mbx2.LitigationHoldOwner
    DeletedItemsRetentionDays  = $mbx2.DeletedItemsRetention.Days
    SingleItemRecoveryEnabled  = $mbx2.SingleItemRecoveryEnabled
    ArchiveStatus              = $mbx2.ArchiveStatus
    AuditEnabled               = $mbx2.AuditEnabled
    AuditLogAgeLimitDays       = $mbx2.AuditLogAgeLimit.Days
    TotalItemSize              = $stats.TotalItemSize
    ItemCount                  = $stats.ItemCount
  } | Format-List
} catch {
  Write-Warning "Verification step encountered an issue: $_"
}

Write-Host "`nDone." -ForegroundColor Green
Stop-Transcript | Out-Null