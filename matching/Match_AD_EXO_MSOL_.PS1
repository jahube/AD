$Cloud =""
$EXORecipients=""
$Mailusers=""
$cloud_filter2 = $null
$Recipient_filter2= $null
$Mbxs2= $null

[System.Collections.ArrayList]$Recipient_filter2 = [System.Collections.ArrayList]$Recipient_filter |where { $_.RCP_RecipientTypeDetails -eq "MailUser" -or $_.RecipientType -eq "UserMailbox"}
[System.Collections.ArrayList]$Recipient_filter = [System.Collections.ArrayList]$Recipient_filter2

 [System.Collections.ArrayList]$cloud_filter2  = [System.Collections.ArrayList]$cloud_filter
 [System.Collections.ArrayList]$cloud_filter  = [System.Collections.ArrayList]$cloud_filter2

 [System.Collections.ArrayList]$Mbxs2 =  [System.Collections.ArrayList]$Mbxs
 [System.Collections.ArrayList]$Mbxs =  [System.Collections.ArrayList]$Mbxs2


[System.Collections.ArrayList]$data =@()

$ADUser_duplicates = @()
$i= 1
$count =  $Recipient_filter.count
foreach ($Recipientinfo in $Recipient_filter) {

Write-Progress -Activity "running" -Status $i -PercentComplete (($i/$count)*100) ; $i++
$MSOLuser = $MSOL_Filter.where({ $Recipientinfo.ExternalDirectoryObjectId -eq  $_.ObjectId })

IF ($Recipientinfo.RecipientType -eq "UserMailbox") { $Clouduser = $cloud_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) 
} ELSEIF ($Recipientinfo.RecipientType -eq "MailUser") { $Clouduser = $Mailuser_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) }

#$ADUser =  $Mbxs.where({ $Recipientinfo.exchangeguid.Guid -eq  $_.msExchMailboxGuid })

#$ADUser = $Mbxs.where({ $_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -and $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid })

#$ADUser =  $Mbxs.where({ "$($Recipientinfo.PrimarySMTPAddress)" -eq  "$($_.mail)" })

$ADUser = $Mbxs.where({ ($_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -or $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid) -and "$($Recipientinfo.PrimarySMTPAddress)" -eq  "$($_.mail)" })

IF($ADUser){ $weakmatch = $false  } Else { $weakmatch = $true ;

$ADUser = $mbxs.where({($_.msExchMailboxGuid -eq $Recipientinfo.ExchangeGuid.guid ) -or ($Recipientinfo.userprincipalname -eq $_.userprincipalname) -or ($Recipientinfo.PrimarySMTPAddress -eq $_.PrimarySMTPAddress) -or ($_.mail -eq $Recipientinfo.PrimarySMTPAddress)})

} 

$duplicates = "none"

IF ($ADUser.count -gt 1) { 

IF ($ADUser.count -gt 2) { $ADUser_duplicates += $ADUser[2] }

$duplicates = $ADUser.count ; $ADUser_duplicates += $ADUser[0] ;  $ADUser_duplicates += $ADUser[1] 

$ADUser = $ADUser[0]

}

$item = New-Object -TypeName PSObject

$item | Add-Member -MemberType NoteProperty -Name RCP_RecipientType -Value $Recipientinfo.RecipientType
$item | Add-Member -MemberType NoteProperty -Name RCP_RecipientTypeDetails -Value $Recipientinfo.RecipientTypeDetails
$item | Add-Member -MemberType NoteProperty -Name Weakmatch -Value $weakmatch
$item | Add-Member -MemberType NoteProperty -Name DuplicateMatch -Value $duplicates

$item | Add-Member -MemberType NoteProperty -Name EXO_Alias -Value $Clouduser.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name EXO_UserPrincipalName -Value $Clouduser.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name EXO_PrimarySmtpAddress -Value $Clouduser.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name EXO_DisplayName -Value $Clouduser.DisplayName
$item | Add-Member -MemberType NoteProperty -Name EXO_Name -Value $Clouduser.Name
$item | Add-Member -MemberType NoteProperty -Name EXO_MicrosoftOnlineServicesID -Value $Clouduser.MicrosoftOnlineServicesID
$item | Add-Member -MemberType NoteProperty -Name EXO_Guid -Value $Clouduser.Guid
$item | Add-Member -MemberType NoteProperty -Name EXO_ExchangeGuid -Value $Clouduser.ExchangeGuid
$item | Add-Member -MemberType NoteProperty -Name EXO_ExchangeObjectId -Value $Clouduser.ExchangeObjectId
$item | Add-Member -MemberType NoteProperty -Name EXO_DistinguishedName -Value $Clouduser.DistinguishedName
$item | Add-Member -MemberType NoteProperty -Name EXO_EmailAddresses -Value $Clouduser.EmailAddresses
$item | Add-Member -MemberType NoteProperty -Name EXO_CustomAttribute4 -Value $Clouduser.CustomAttribute4
$item | Add-Member -MemberType NoteProperty -Name EXO_AccountDisabled -Value $Clouduser.AccountDisabled
$item | Add-Member -MemberType NoteProperty -Name EXO_SKUAssigned -Value $Clouduser.SKUAssigned
$item | Add-Member -MemberType NoteProperty -Name EXO_TargetProxyAddress -Value $Clouduser.TargetProxyAddress_EXO
$item | Add-Member -MemberType NoteProperty -Name EXO_Onmicrosoft -Value $Clouduser.Onmicrosoft_EXO
$item | Add-Member -MemberType NoteProperty -Name EXO_RemoteRecipientType -Value $Clouduser.RemoteRecipientType
$item | Add-Member -MemberType NoteProperty -Name EXO_RecipientTypeDetails -Value $Clouduser.RecipientTypeDetails
$item | Add-Member -MemberType NoteProperty -Name EXO_RecipientType -Value $Clouduser.RecipientType

IF($Recipientinfo.RecipientType -eq "MailUser") {
$item | Add-Member -MemberType NoteProperty -Name EXO_Externalemailaddress -Value $($Clouduser.Externalemailaddress -ireplace "SMTP:")
} ELSE { $item | Add-Member -MemberType NoteProperty -Name EXO_Externalemailaddress -Value "KEIN" }

$item | Add-Member -MemberType NoteProperty -Name MSOL_userprincipalname -Value $MSOLuser.userprincipalname
$item | Add-Member -MemberType NoteProperty -Name MSOL_PrimarySMTPAddress -Value $MSOLuser.PrimarySMTPAddress
$item | Add-Member -MemberType NoteProperty -Name MSOL_TargetProxyAddress -Value $MSOLuser.TargetProxyAddress_MSOL
$item | Add-Member -MemberType NoteProperty -Name MSOL_immutableID -Value $MSOLuser.immutableID
$item | Add-Member -MemberType NoteProperty -Name MSOL_Immutable_converted -Value $MSOLuser.Immutable_converted
$item | Add-Member -MemberType NoteProperty -Name MSOL_MSExchRecipientTypeDetails -Value $MSOLuser.MSExchRecipientTypeDetails
$item | Add-Member -MemberType NoteProperty -Name MSOL_DisplayName -Value $MSOLuser.DisplayName
$item | Add-Member -MemberType NoteProperty -Name MSOL_FirstName -Value $MSOLuser.FirstName
$item | Add-Member -MemberType NoteProperty -Name MSOL_LastName -Value $MSOLuser.LastName
$item | Add-Member -MemberType NoteProperty -Name MSOL_ObjectId -Value $MSOLuser.ObjectId
$item | Add-Member -MemberType NoteProperty -Name MSOL_LiveId -Value $MSOLuser.LiveId
$item | Add-Member -MemberType NoteProperty -Name MSOL_Onmicrosoft -Value $MSOLuser.Onmicrosoft_MSOL

IF ($ADUser) { 

$item | Add-Member -MemberType NoteProperty -Name AD_UserPrincipalName -Value $ADUser.UserPrincipalName
$item | Add-Member -MemberType NoteProperty -Name AD_PrimarySmtpAddress -Value $ADUser.PrimarySmtpAddress
$item | Add-Member -MemberType NoteProperty -Name AD_samaccountname -Value $ADUser.samaccountname
$item | Add-Member -MemberType NoteProperty -Name AD_msExchMailboxGuid -Value $ADUser.msExchMailboxGuid
$item | Add-Member -MemberType NoteProperty -Name AD_ObjectGUID -Value $ADUser.ObjectGUID
$item | Add-Member -MemberType NoteProperty -Name AD_mail -Value $ADUser.mail
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRemoteRecipientType -Value $ADUser.msExchRemoteRecipientType
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRecipientDisplayType -Value $ADUser.msExchRecipientDisplayType
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRecipientTypeDetails -Value $ADUser.msExchRecipientTypeDetails
$item | Add-Member -MemberType NoteProperty -Name AD_extensionattribute4 -Value $ADUser.extensionattribute4
$item | Add-Member -MemberType NoteProperty -Name AD_enabled -Value $ADUser.enabled
$item | Add-Member -MemberType NoteProperty -Name AD_lastlogondate -Value $ADUser.lastlogondate
$item | Add-Member -MemberType NoteProperty -Name AD_TargetAddress -Value $ADUser.TargetAddress
$item | Add-Member -MemberType NoteProperty -Name AD_TargetProxyAddress -Value $ADUser.TargetProxyAddress
$item | Add-Member -MemberType NoteProperty -Name AD_proxyaddresses -Value $ADUser.proxyaddresses
$item | Add-Member -MemberType NoteProperty -Name AD_OU -Value $ADUser."Top-Level_OU"
$item | Add-Member -MemberType NoteProperty -Name AD_SMTP_Domain -Value $ADUser.SMTP_Domain
$item | Add-Member -MemberType NoteProperty -Name AD_UPN_Domain -Value $ADUser.UPN_Domain
$item | Add-Member -MemberType NoteProperty -Name AD_MAIL_Domain -Value $ADUser.MAIL_Domain
$item | Add-Member -MemberType NoteProperty -Name AD_SMTP_Alias -Value $ADUser.SMTP_Alias
$item | Add-Member -MemberType NoteProperty -Name AD_UPN_Alias -Value $ADUser.UPN_Alias

} ELSE {

$item | Add-Member -MemberType NoteProperty -Name AD_UserPrincipalName -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_PrimarySmtpAddress -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_samaccountname -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_msExchMailboxGuid -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_ObjectGUID -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_mail -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRemoteRecipientType -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRecipientDisplayType -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_msExchRecipientTypeDetails -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_extensionattribute4 -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_enabled -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_lastlogondate -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_TargetAddress -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_TargetProxyAddress -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_proxyaddresses -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_OU -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_SMTP_Domain -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_UPN_Domain -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_MAIL_Domain -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_SMTP_Alias -Value "KEIN_ADUSER"
$item | Add-Member -MemberType NoteProperty -Name AD_UPN_Alias -Value "KEIN_ADUSER"

}

$data += $item

}

$data | export-Clixml $path\data_ELKW_NEU7.xml -Force

$data | Export-Csv $path\data_ELKW_NEU7.csv -Encoding UTF8 -Delimiter ";" -Force  -NTI

$data.count

$ADUser_duplicates| Export-Csv $path\ADUser_data_ELKW_NEU7.csv -Encoding UTF8 -Delimiter ";" -NTI -Force

$ADUser_duplicates| export-Clixml $path\ADUser_data_ELKW_NEU7.xml

$MSOLuser.TargetProxyAddress_MSOL

$MSOLuser.TargetProxyAddress_MSOL

2.count

$data2 = $data.where({ $_.RCP_RecipientType -eq "UserMailbox" -and $_.MSOL_PrimarySMTPAddress -notmatch "elk-wue.de$" })

$mismatch = $data2.where({ "$($_.AD_TargetAddress)" -ne "KEIN_ADUSER" -and "$($_.AD_TargetAddress -replace "SMTP:")" -ne "$($_.MSOL_TargetProxyAddress)" })

$mismatch = $data.where({ $_.RCP_RecipientType -eq "UserMailbox" -and "$($_.AD_TargetAddress)" -ne "KEIN_ADUSER" -and "$($_.AD_TargetAddress)" -ne "$('SMTP:' + $_.EXO_TargetProxyAddress)" })

$mismatch |ft AD_TargetAddress,MSOL_TargetProxyAddress,AD_OU,EXO_TargetProxyAddress -AutoSize
$mismatch.count

$fix = $mismatch.where({ $_.AD_OU -notmatch "Gast" -and $_.AD_OU -notmatch "krbtgt" -and $_.samaccountname -notmatch "Gast" -and  $_.samaccountname -notmatch "krbtgt" -and $_.RCP_RecipientType -eq "UserMailbox" -and ("$($_.AD_TargetAddress)" -replace 'SMTP:') -notin $_.EXO_EmailAddresses })

$fix |ft ad_mail,AD_TargetAddress,AD_OU,EXO_TargetProxyAddress,EXO_EmailAddresses -AutoSize
$r_duplicates = $ADUser_duplicates.where({ $_.AD_samaccountname -ne "Gast" -and  $_.AD_samaccountname -ne "krbtgt"  })
Gast

$r_duplicates  |ft ad_mail,AD_TargetAddress,AD_OU,EXO_TargetProxyAddress,EXO_EmailAddresses -AutoSize

$match7 = $data2.where({ $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })
$no_match7 = $data2.where({ $_.EXO_ExchangeGuid -ne $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })

$match9 = $data2.where({ $_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -and $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })
$no_match9 = $data2.where({ $_.MSOL_Immutable_converted -ne $_.AD_ObjectGUID -and $_.EXO_ExchangeGuid -ne $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })

$match7.count
$no_match7 | ft ad_samaccountname,ad_ou,EXO_ExchangeGuid,AD_msExchMailboxGuid

$no_match7.count


$match8.count
$no_match8 | ft ad_samaccountname,ad_ou,EXO_ExchangeGuid,AD_msExchMailboxGuid

$no_match8.count

$match9.count
$no_match9 | ft ad_samaccountname,ad_ou,EXO_ExchangeGuid,AD_msExchMailboxGuid

$no_match9.count

$half_match = $data2.where({ (($_.MSOL_Immutable_converted -ne $_.AD_ObjectGUID -or $_.EXO_ExchangeGuid -ne $_.AD_msExchMailboxGuid) -and ($_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -or $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid ) )-and $_.AD_OU -notmatch "gast" })
$half_match.count

$match11 = $data2.where({ $_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -and $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })

$no_match11 = $data2.where({ $_.MSOL_Immutable_converted -ne $_.AD_ObjectGUID -and $_.EXO_ExchangeGuid -ne $_.AD_msExchMailboxGuid -and $_.AD_OU -notmatch "gast" })

$double = $data2.where({($_.MSOL_Immutable_converted -eq $_.AD_ObjectGUID -or $_.EXO_ExchangeGuid -eq $_.AD_msExchMailboxGuid) -and $_.EXO_PrimarySmtpAddress -eq $_.AD_mail  })

$double.count

$no_match11.count

$match11.count