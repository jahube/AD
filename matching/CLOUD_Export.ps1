
$onmicrosoft = "elkw.onmicrosoft.com"
$routingdomain = "elkw.mail.onmicrosoft.com"

$MSOLUSers = get-msoluser -all

$MSOL_Filter = $MSOLUSers | select userprincipalname,immutableID,@{N="Immutable_converted" ;E={ [Guid]([Convert]::FromBase64String($_.ImmutableId)) }},MSExchRecipientTypeDetails,DisplayName,FirstName,LastName,ObjectId,LiveId,@{N="Onmicrosoft_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }},@{N="TargetProxyAddress_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="proxyaddresses" ;E={ ($_.proxyaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:" }},@{N="proxyaddresses_join" ;E={ (($_.proxyaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:") -join '|' }}


$Cloud = Get-EXOMailbox -ResultSize unlimited -Properties PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned

$cloud_filter = $Cloud  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,@{N="emailaddresses" ;E={ ($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:" }},@{N="emailaddresses_join" ;E={ (($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:") -join '|' }},CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}


$EXORecipients = Get-EXORecipient -ResultSize unlimited -Properties PrimarySmtpAddress,Alias,ExternalDirectoryObjectId,ExternalEmailAddress,RecipientTypeDetails,RecipientType,Name,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,SKUAssigned

$Recipient_filter = $EXORecipients  | select PrimarySmtpAddress,Alias,ExternalDirectoryObjectId,@{N="ExternalEmailAddress" ;E={ $_.ExternalEmailAddress -replace "smtp:" }},RecipientTypeDetails,RecipientType,Name,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,@{N="emailaddresses" ;E={ ($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:" }},@{N="emailaddresses_join" ;E={ (($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:") -join '|' }},CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -ireplace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }},@{N="X500_EXO" ;E={ ($_.emailaddresses | where { $_ -match "^x500" })}},@{N="SIP_EXO" ;E={ ($_.emailaddresses | where { $_ -match "^SIP" })}},@{N="SPO_EXO" ;E={ ($_.emailaddresses | where { $_ -match "^SPO" })}}


$Mailusers = Get-MailUser -ResultSize unlimited

$Mailuser_filter = $Mailusers  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,ExternalEmailAddress,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,@{N="emailaddresses" ;E={ ($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:" }},@{N="emailaddresses_join" ;E={ (($_.emailaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:") -join '|' }},CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

