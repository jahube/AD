
$data =@()

$cloud_duplicates =@()

$cloud_notfound_data =@()

foreach ($ADUser in $mbxs) {

$MSOLuser = $null
$hardmatch = $null
$softmatch = $null
$UPNmatch = $null

$ObjectGUID = $ADUser.ObjectGUID
$MSOLuser = $MSOL_Filter.where({ $ObjectGUID -eq $_.Immutable_converted })

IF ($MSOLuser) { 
$Hardmatch = $true
$softmatch = $false
$UPNmatch = $false
} ELSE { 
$MSOLuser = $MSOL_Filter.where({ $ADUser.PrimarySMTPAddress -eq $_.PrimarySMTPAddress })
IF ($MSOLuser) {
$Hardmatch = $false
$softmatch = $true
$UPNmatch = $false
} ELSE {
$MSOLuser = $MSOL_Filter.where({ $ADUser.userprincipalname -eq $_.userprincipalname })
$Hardmatch = $false
$softmatch = $false
$UPNmatch = $true
       }
}

$duplicatematch = $null
$cloud_notfound = $null
$EXO_type = $null

$Clouduser = $cloud_filter.where({($MSOLuser.ObjectId -eq $_.ExternalDirectoryObjectId) -or ($ADUser.msExchMailboxGuid -eq $_.ExchangeGuid) -or (($ADUser.userprincipalname -eq $_.userprincipalname) -and ($ADUser.PrimarySMTPAddress -eq $_.PrimarySMTPAddress -or $ADUser.mail -eq $_.PrimarySMTPAddress )) })
if ($Clouduser.count -ne 0) { $EXO_type = "CLOUD_Mailbox" }

if ($Clouduser.count -eq 0) {

$Clouduser = $Mailuser_filter.where({($MSOLuser.ObjectId -eq $_.ExternalDirectoryObjectId) -or ($ADUser.msExchMailboxGuid -eq $_.ExchangeGuid) -or (($ADUser.userprincipalname -eq $_.userprincipalname) -and ($ADUser.PrimarySMTPAddress -eq $_.PrimarySMTPAddress -or $ADUser.mail -eq $_.PrimarySMTPAddress )) })

if ($Clouduser.count -ne 0) { $EXO_type = "ONPREM_Mailbox" }

}


$item = New-Object -TypeName PSObject
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

IF ($Clouduser.count -eq 1) {

$item | Add-Member -MemberType NoteProperty -Name EXO_type -Value $EXO_type
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

}

IF ($Clouduser.count -gt 1) {
$duplicatematch = $true
$item | Add-Member -MemberType NoteProperty -Name EXO_duplicatematch -Value $duplicatematch
}

IF ($Clouduser.count -eq "0") {
$cloud_notfound = $true
$item | Add-Member -MemberType NoteProperty -Name EXO_cloud_notfound -Value $cloud_notfound
}

IF ($MSOLuser.count -eq 1) {

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

}

IF (($MSOLuser.count -eq 1) -and ($Clouduser.count -eq "1")) {

$data += $item

} ELSE {

IF (($MSOLuser.count -gt 1) -or ($Clouduser.count -eq "1")) {

$cloud_duplicates  += $item

}

IF ($Clouduser.count -eq 0) {

$cloud_notfound_data += $item

}

       }

}
