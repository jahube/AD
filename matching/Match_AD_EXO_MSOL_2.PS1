
#[System.Collections.ArrayList]$Recipient_filter = [System.Collections.ArrayList]$Recipient_filter2 |where { $_.RCP_RecipientTypeDetails -eq "MailUser" -or $_.RecipientType -eq "UserMailbox"}


 #[System.Collections.ArrayList]$cloud_filter  = [System.Collections.ArrayList]$cloud_filter2
 #[System.Collections.ArrayList]$Mbxs =  [System.Collections.ArrayList]$Mbxs2

[System.Collections.ArrayList]$data =@()

$ADUser_duplicates = @()

$i= 1
$count =  $Recipient_filter.count
foreach ($Recipientinfo in $Recipient_filter) {
Write-Progress -Activity "running" -Status $i -PercentComplete (($i/$count)*100) ; $i++
#$MSOLuser = $MSOL_Filter.where({ [guid]$Recipientinfo.ExternalDirectoryObjectId -match [guid]$_.ObjectId.Guid })
$MSOLuser = $MSOL_Filter.where({ $Recipientinfo.userprincipalname -match $_.userprincipalname })

#$MSOLuser = $MSOL_Filter.where({ [guid]$Recipientinfo.ExternalDirectoryObjectId -eq [guid]$_.ObjectId.Guid })


IF ($Recipientinfo.RecipientType -eq "UserMailbox") { $Clouduser = $cloud_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) 

} ELSEIF ($Recipientinfo.RecipientType -eq "MailUser") { $Clouduser = $Mailuser_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) }

#IF ($Recipientinfo.RecipientType -eq "UserMailbox") { $Clouduser = $cloud_filter.where({$Recipientinfo.userprincipalname -eq $_.userprincipalname }) }

#IF ($Recipientinfo.RecipientType -eq "MailUser") { $Clouduser = $Mailuser_filter.where({$Recipientinfo.userprincipalname -eq $_.userprincipalname }) }

#$ADUser =  $Mbxs.where({ $Recipientinfo.exchangeguid.Guid -eq  $_.msExchMailboxGuid })

#IF ($Recipientinfo.RCP_RecipientTypeDetails -eq "MailUser" -or $Recipientinfo.RecipientType -eq "UserMailbox") {
$ADUser =  $Mbxs.where({ "$($Recipientinfo.PrimarySMTPAddress)" -eq  "$($_.mail)" })
#}

#IF ($Recipientinfo.RCP_RecipientTypeDetails -eq "MailUser" -or $Recipientinfo.RecipientType -eq "UserMailbox") {
#$ADUser =  $Mbxs.where({ "$($Recipientinfo.PrimarySMTPAddress)" -eq  "$($_.mail)" })
#}



#IF(!($ADUser)) {
#$ADUser = $mbxs.where({($_.msExchMailboxGuid -eq $Recipientinfo.ExchangeGuid.guid ) -or ($Recipientinfo.userprincipalname -eq $_.userprincipalname) -or ($Recipientinfo.PrimarySMTPAddress -eq $_.PrimarySMTPAddress) -or ($_.mail -eq $Recipientinfo.PrimarySMTPAddress)})
#}

$duplicates = "none"

#IF ($ADUser.count -gt 2) { $ADUser_duplicates += $ADUser[2] }
#IF ($ADUser.count -gt 1) { $duplicates = $ADUser.count ; $ADUser_duplicates += $ADUser[0] ;  $ADUser_duplicates += $ADUser[1] 
#$ADUser = $ADUser[0]
#}

IF ($ADUser.count -gt 1) {  $ADUser_duplicates +=$ADUser ; $ADUser = $ADUser[0] }

IF ($ADUser) { 

$item = [PSCustomObject]@{RCP_RecipientType = $Recipientinfo.RecipientType
RCP_RecipientTypeDetails = $Recipientinfo.RecipientTypeDetails
EXO_Alias = $Clouduser.PrimarySmtpAddress
EXO_UserPrincipalName = $Clouduser.PrimarySmtpAddress
EXO_PrimarySmtpAddress = $Clouduser.PrimarySmtpAddress
EXO_DisplayName = $Clouduser.DisplayName
EXO_Name = $Clouduser.Name
EXO_MicrosoftOnlineServicesID = $Clouduser.MicrosoftOnlineServicesID
EXO_Guid = $Clouduser.Guid
EXO_ExchangeGuid = $Clouduser.ExchangeGuid
EXO_ExchangeObjectId = $Clouduser.ExchangeObjectId
EXO_DistinguishedName = $Clouduser.DistinguishedName
EXO_EmailAddresses = $Clouduser.EmailAddresses
EXO_CustomAttribute4 = $Clouduser.CustomAttribute4
EXO_AccountDisabled = $Clouduser.AccountDisabled
EXO_SKUAssigned = $Clouduser.SKUAssigned
EXO_TargetProxyAddress = $Clouduser.TargetProxyAddress_EXO
EXO_Onmicrosoft = $Clouduser.Onmicrosoft_EXO
EXO_RemoteRecipientType = $Clouduser.RemoteRecipientType
EXO_RecipientTypeDetails = $Clouduser.RecipientTypeDetails
EXO_RecipientType = $Clouduser.RecipientType

MSOL_userprincipalname = $MSOLuser.userprincipalname
MSOL_PrimarySMTPAddress = $MSOLuser.PrimarySMTPAddress
MSOL_TargetProxyAddress = $MSOLuser.TargetProxyAddress_MSOL
MSOL_immutableID = $MSOLuser.immutableID
MSOL_Immutable_converted = $MSOLuser.Immutable_converted
MSOL_MSExchRecipientTypeDetails = $MSOLuser.MSExchRecipientTypeDetails
MSOL_DisplayName = $MSOLuser.DisplayName
MSOL_FirstName = $MSOLuser.FirstName
MSOL_LastName = $MSOLuser.LastName
MSOL_ObjectId = $MSOLuser.ObjectId
MSOL_LiveId = $MSOLuser.LiveId
MSOL_Onmicrosoft = $MSOLuser.Onmicrosoft_MSOL

AD_UserPrincipalName = $ADUser.UserPrincipalName
AD_PrimarySmtpAddress = $ADUser.PrimarySmtpAddress
AD_samaccountname = $ADUser.samaccountname
AD_msExchMailboxGuid = $ADUser.msExchMailboxGuid
AD_ObjectGUID = $ADUser.ObjectGUID
AD_mail = $ADUser.mail
AD_msExchRemoteRecipientType = $ADUser.msExchRemoteRecipientType
AD_msExchRecipientDisplayType = $ADUser.msExchRecipientDisplayType
AD_msExchRecipientTypeDetails = $ADUser.msExchRecipientTypeDetails
AD_extensionattribute4 = $ADUser.extensionattribute4
AD_enabled = $ADUser.enabled
AD_lastlogondate = $ADUser.lastlogondate
AD_TargetAddress = $ADUser.TargetAddress
AD_TargetProxyAddress = $ADUser.TargetProxyAddress
AD_proxyaddresses = $ADUser.proxyaddresses
}

} ELSE {

$item = [PSCustomObject]@{RCP_RecipientType = $Recipientinfo.RecipientType
RCP_RecipientTypeDetails = $Recipientinfo.RecipientTypeDetails
EXO_Alias = $Clouduser.PrimarySmtpAddress
EXO_UserPrincipalName = $Clouduser.PrimarySmtpAddress
EXO_PrimarySmtpAddress = $Clouduser.PrimarySmtpAddress
EXO_DisplayName = $Clouduser.DisplayName
EXO_Name = $Clouduser.Name
EXO_MicrosoftOnlineServicesID = $Clouduser.MicrosoftOnlineServicesID
EXO_Guid = $Clouduser.Guid
EXO_ExchangeGuid = $Clouduser.ExchangeGuid
EXO_ExchangeObjectId = $Clouduser.ExchangeObjectId
EXO_DistinguishedName = $Clouduser.DistinguishedName
EXO_EmailAddresses = $Clouduser.EmailAddresses
EXO_CustomAttribute4 = $Clouduser.CustomAttribute4
EXO_AccountDisabled = $Clouduser.AccountDisabled
EXO_SKUAssigned = $Clouduser.SKUAssigned
EXO_TargetProxyAddress = $Clouduser.TargetProxyAddress_EXO
EXO_Onmicrosoft = $Clouduser.Onmicrosoft_EXO
EXO_RemoteRecipientType = $Clouduser.RemoteRecipientType
EXO_RecipientTypeDetails = $Clouduser.RecipientTypeDetails
EXO_RecipientType = $Clouduser.RecipientType

MSOL_userprincipalname = $MSOLuser.userprincipalname
MSOL_PrimarySMTPAddress = $MSOLuser.PrimarySMTPAddress
MSOL_TargetProxyAddress = $MSOLuser.TargetProxyAddress_MSOL
MSOL_immutableID = $MSOLuser.immutableID
MSOL_Immutable_converted = $MSOLuser.Immutable_converted
MSOL_MSExchRecipientTypeDetails = $MSOLuser.MSExchRecipientTypeDetails
MSOL_DisplayName = $MSOLuser.DisplayName
MSOL_FirstName = $MSOLuser.FirstName
MSOL_LastName = $MSOLuser.LastName
MSOL_ObjectId = $MSOLuser.ObjectId
MSOL_LiveId = $MSOLuser.LiveId
MSOL_Onmicrosoft = $MSOLuser.Onmicrosoft_MSOL

AD_UserPrincipalName = "KEIN_ADUSER"
AD_PrimarySmtpAddress = "KEIN_ADUSER"
AD_samaccountname = "KEIN_ADUSER"
AD_msExchMailboxGuid = "KEIN_ADUSER"
AD_ObjectGUID = "KEIN_ADUSER"
AD_mail = "KEIN_ADUSER"
AD_msExchRemoteRecipientType = "KEIN_ADUSER"
AD_msExchRecipientDisplayType = "KEIN_ADUSER"
AD_msExchRecipientTypeDetails = "KEIN_ADUSER"
AD_extensionattribute4 = "KEIN_ADUSER"
AD_enabled = "KEIN_ADUSER"
AD_lastlogondate = "KEIN_ADUSER"
AD_TargetAddress = "KEIN_ADUSER"
AD_TargetProxyAddress = "KEIN_ADUSER"
AD_proxyaddresses = "KEIN_ADUSER"
}

$item |Export-csv "c:\Temp\ELKW_TEST_flow.csv" -append -encoding utf8 -D ";" -NTI
#$data.add($item)
}


}