$mailuser = Get-MailUser -ResultSize unlimited
$Kontakte = Get-MailContact -ResultSize unlimited
$mbxs = get-mailbox -ResultSize unlimited
$SmtpForwarding = $mbxs | Select DisplayName,userprincipalname,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxandForward | Where-Object {$null -ne $_.ForwardingSMTPAddress -or $null -ne $_.ForwardingAddress }
$SmtpForwarding |ft -AutoSize
$data = @()
foreach ($D in $SmtpForwarding) {
$ForwardingSMTPAddress = ""
IF ($D.ForwardingSMTPAddress -ne "") {
$ForwardingSMTPAddress = $D.ForwardingSMTPAddress
$item = New-Object -TypeName PSCustomObject
$item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $D.DisplayName
$item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $D.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $D.ForwardingAddress
$item | Add-Member -MemberType NoteProperty -Name ForwardingSMTPAddress -Value $ForwardingSMTPAddress
$item | Add-Member -MemberType NoteProperty -Name Type -Value "ForwardingSMTPAddress"
$item | Add-Member -MemberType NoteProperty -Name user -Value "ForwardingSMTPAddress"
$item | Add-Member -MemberType NoteProperty -Name external -Value ExternalEmailAddress
$item | Add-Member -MemberType NoteProperty -Name Hidden -Value sichtbar
}
IF (!($ForwardingSMTPAddress )) {
$muser = $mailuser | where { $D.ForwardingAddress -match $_.alias -or $D.ForwardingAddress -match $_.name  -or $D.ForwardingAddress -match $_.displayname }
}
IF ($muser) {
$item = New-Object -TypeName PSCustomObject
$item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $D.DisplayName
$item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $D.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $D.ForwardingAddress
$item | Add-Member -MemberType NoteProperty -Name ForwardingSMTPAddress -Value $muser.ExternalEmailAddress
$item | Add-Member -MemberType NoteProperty -Name Type -Value "mailuser"
$item | Add-Member -MemberType NoteProperty -Name user -Value $muser.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name external -Value $muser.ExternalEmailAddress
$item | Add-Member -MemberType NoteProperty -Name Hidden -Value $muser.HiddenFromAddressListsEnabled
}
IF (!($muser)) {
$mbx  = $mbxs | where { $D.ForwardingAddress -match $_.alias -or $D.ForwardingAddress -match $_.name  -or $D.ForwardingAddress -match $_.displayname }
}
IF ($mbx) {
$item = New-Object -TypeName PSCustomObject
$item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $D.DisplayName
$item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $D.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $D.ForwardingAddress
$item | Add-Member -MemberType NoteProperty -Name ForwardingSMTPAddress -Value $mbx.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name Type -Value "Mailbox"
$item | Add-Member -MemberType NoteProperty -Name user -Value $mbx.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name Hidden -Value $mbx.HiddenFromAddressListsEnabled
}
IF (!($mbx)) {
$kontakt = $Kontakte  | where { $D.ForwardingAddress -match $_.alias -or $D.ForwardingAddress -match $_.name  -or $D.ForwardingAddress -match $_.displayname }
}
IF ($kontakt) {
$item = New-Object -TypeName PSCustomObject
$item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $D.DisplayName
$item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $D.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $D.ForwardingAddress
$item | Add-Member -MemberType NoteProperty -Name ForwardingSMTPAddress -Value $ForwardingSMTPAddress
$item | Add-Member -MemberType NoteProperty -Name Type -Value "Kontakt"
$item | Add-Member -MemberType NoteProperty -Name user -Value $kontakt.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name external -Value $kontakt.ExternalEmailAddress
$item | Add-Member -MemberType NoteProperty -Name Hidden -Value $kontakt.HiddenFromAddressListsEnabled
}
IF (!($kontakt) -and !($mbx) -and !($muser) -and !($ForwardingSMTPAddress)){
$item = New-Object -TypeName PSCustomObject
$item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $D.DisplayName
$item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $D.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $D.ForwardingAddress
$item | Add-Member -MemberType NoteProperty -Name ForwardingSMTPAddress -Value "unbekannt"
$item | Add-Member -MemberType NoteProperty -Name Type -Value "unbekannt"
$item | Add-Member -MemberType NoteProperty -Name user -Value "unbekannt"
$item | Add-Member -MemberType NoteProperty -Name external -Value "unbekannt"
$item | Add-Member -MemberType NoteProperty -Name Hidden -Value "unbekannt"
}
$data += $item
}
$data | ConvertTo-Html | Out-File "C:\temp\SmtpForwarding.html" -force
$data | export-Csv C:\Temp\SmtpForwarding.csv -Encoding UTF8 -Delimiter ";" -NTI
  $UserInboxRules = @()
  $UserDelegates = @()
foreach ($User in $mbxs) {
$UserInboxRules += ( Invoke-Command { Get-InboxRule -Mailbox $User.UserPrincipalname -ErrorAction SilentlyContinue } | Select-Object -Property Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | Where-Object {($null -ne $_.ForwardTo) -or ($null -ne $_.ForwardAsAttachmentTo) -or ($null -ne $_.RedirectsTo)})
$UserDelegates += Get-MailboxPermission -Identity $User.UserPrincipalName | Where-Object {($_.IsInherited -ne 'True') -and ($_.User -notlike '*SELF*')}
}
$UserInboxRules  | ConvertTo-Html | Out-File "C:\temp\UserInboxRules.html" -force
$UserInboxRules | export-Csv C:\Temp\UserInboxRules.csv -Encoding UTF8 -Delimiter ";" -NTI
$UserDelegates | ConvertTo-Html | Out-File "C:\temp\UserDelegates.html" -force
$UserDelegates | export-Csv C:\Temp\UserDelegates.csv -Encoding UTF8 -Delimiter ";" -NTI
$path2 = "C:\temp\test.html"
$SmtpForwarding = [pscustomobject] @{
"Group Name" = "Authenticated Users"
"Members" = "User1", "User2", "User3"
}