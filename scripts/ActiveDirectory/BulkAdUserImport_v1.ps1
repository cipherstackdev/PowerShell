 Import-Module Ac*

Import-Module ActiveDirectory
$rows = Import-Csv C:\dir\bulk-user-import.csv

foreach ($r in $rows) {
    $upn = ($r.userprincipalname).Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Warning "Skipping $($r.samAccountName): missing userprincipalname"
        continue
    }

    New-ADUser `
      -GivenName $r.givenName.Trim() `
      -Surname $r.surname.Trim() `
      -Name $r.name.Trim() `
      -DisplayName $r.displayname.Trim() `
      -SamAccountName $r.samAccountName.Trim() `
      -UserPrincipalName $upn `
      -EmailAddress $upn `
      -Path $r.path.Trim() `
      -AccountPassword (ConvertTo-SecureString $r.password -AsPlainText -Force) `
      -Enabled $true `
      -PasswordNeverExpires $true `
      -Office $r.office `
      -Title $r.title
}