$fehler | ft defaultpublicfoldermailbox ,customattribute4,recipienttypedetails,primarysmtpaddress -AutoSize

$fehler | % { set-mailbox $_.Primarysmtpaddress -defaultpublicfoldermailbox PFEND  } 

$path = "C:\temp"
$file = "Alle_CLOUD_Benutzer_export.CSV"

$data = import-csv "$path\$file" -delimiter ";" -encoding UtF8

$mbxs = foreach ($user in $data) {

if ($user.mail -match "domain.de") {

$mail = $user.mail

invoke-command { Get-EXOMailbox $mail -Properties defaultpublicfoldermailbox,customattribute4,recipienttypedetails,primarysmtpaddress,MessageCopyForSentAsEnabled,MessageCopyForSendOnBehalfEnabled } -ArgumentList $mail

}

}

$mbxs.count

$FPPFEND_err = $mbxs | where { $_.defaultpublicfoldermailbox -ne "PFEND"  }

$FPPFEND_err.count


$FPPFEND_err = $mbxs | where { $_.defaultpublicfoldermailbox -ne "PFEND"  }

$FPPFEND_err.count

$FP_notshared = $mbxs | where { $_.MessageCopyForSentAsEnabled -ne $true -or  $_.MessageCopyForSendOnBehalfEnabled -ne $true  }

$FP_notshared | % { set-mailbox $_.distinguishedname -MessageCopyforSentAsEnabled $true -MessageCopyforSendOnBehalfEnabled $true }


$mbxs = Get-MailboxPermission

$permissions = $mbxs | get-mailboxpermission| where {$_.isinherited -eq $false -and $_.user -notlike "NT*Authority*" }

$permissions | select *,@{n = "accessrights"; E={ $_.accessrights -split ', ' -join '|' }} -ExcludeProperty accessrights | export-csv C:\Temp\Mailbox_Berechtigungen.csv -Delimiter ";" -Encoding UTF8 -Force-NTI



$FPPFEND_err = $mbxs | where { $_.recipienttypedetails -ne "Sharedmailbox" -and $_.customattribute4 -notmatch "AP-Typ"  }

$FPPFEND_err | select alias,userprincipalname,exchangeguid,defaultpublicfoldermailbox,customattribute4,recipienttypedetails,primarysmtpaddress | export-csv "$path\Missing_license.csv" -Delimiter ";" -Encoding UTF8 -NTI

Connect-MsolService

$AP_Update = foreach ($user in $FPPFEND_err.userprincipalname) { Get-EXOMailbox $user -Properties alias,userprincipalname,exchangeguid,defaultpublicfoldermailbox,customattribute4,recipienttypedetails,primarysmtpaddress }

$AP_Update | ft userprincipalname,customattribute4


$MSOL = foreach ($user in $FPPFEND_err.userprincipalname) { Get-MsolUser -UserPrincipalName $user }

$missing = $MSOL |where { $_.Licenses.AccountSkuId -notmatch "enterprise" }