
$data =@()

$ADUser_duplicates = @()

foreach ($Recipientinfo in $Recipient_filter) {

$MSOLuser = $MSOL_Filter.where({ $Recipientinfo.ExternalDirectoryObjectId -eq  $_.ObjectId })

IF ($Recipientinfo.RecipientType -eq "UserMailbox") { $Clouduser = $cloud_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) }

IF ($Recipientinfo.RecipientType -eq "MailUser") { $Clouduser = $Mailuser_filter.where({$Recipientinfo.ExternalDirectoryObjectId -eq $_.ExternalDirectoryObjectId }) }

#$ADUser =  $Mbxs.where({ $Recipientinfo.exchangeguid.Guid -eq  $_.msExchMailboxGuid })

$ADUser = $mbxs.where({($_.msExchMailboxGuid -eq $Recipientinfo.ExchangeGuid.guid ) -or ($Recipientinfo.userprincipalname -eq $_.userprincipalname) -or ($Recipientinfo.PrimarySMTPAddress -eq $_.PrimarySMTPAddress) -or ($_.mail -eq $Recipientinfo.PrimarySMTPAddress)})

$duplicates = "none"

IF ($ADUser.count -gt 2) { $ADUser_duplicates += $ADUser[2] }
IF ($ADUser.count -gt 1) { $duplicates = $ADUser.count ; $ADUser_duplicates += $ADUser[0] ;  $ADUser_duplicates += $ADUser[1] 

$ADUser = $ADUser[0]

}

$item = New-Object -TypeName PSObject

$item | Add-Member -MemberType NoteProperty -Name RCP_RecipientType -Value $Recipientinfo.RecipientType
$item | Add-Member -MemberType NoteProperty -Name RCP_RecipientTypeDetails -Value $Recipientinfo.RecipientTypeDetails

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

$item | Add-Member -MemberType NoteProperty -Name Hardmatch -Value $Hardmatch
$item | Add-Member -MemberType NoteProperty -Name softmatch -Value $softmatch
$item | Add-Member -MemberType NoteProperty -Name UPNmatch -Value $UPNmatch

}

$data += $item


}

