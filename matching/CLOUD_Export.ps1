$MSOLUSers = get-msoluser -all
$MSOL_Filter = $MSOLUSers | select userprincipalname,immutableID,@{N="Immutable_converted" ;E={ [Guid]([Convert]::FromBase64String($_.ImmutableId)) }},MSExchRecipientTypeDetails,DisplayName,FirstName,LastName,ObjectId,LiveId,@{N="Onmicrosoft_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }},@{N="TargetProxyAddress_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }}

$MSOL_Filter.count

$EXORecipients = Get-EXORecipient -ResultSize unlimited -PropertySets all
$Recipient_filter = $EXORecipients  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$Mailusers = Get-MailUser -ResultSize unlimited
$Mailuser_filter = $Mailusers  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$Cloud = Get-EXOMailbox -ResultSize unlimited -PropertySets all
$cloud_filter = $Cloud  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$cloud_filter.count
