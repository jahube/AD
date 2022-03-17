$Recipient_filter = $EXORecipients  | select PrimarySmtpAddress,ExternalEmailAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$MSOLUSers = get-msoluser -all

$MSOL_Filter = $MSOLUSers | select userprincipalname,immutableID,@{N="Immutable_converted" ;E={ [Guid]([Convert]::FromBase64String($_.ImmutableId)) }},MSExchRecipientTypeDetails,DisplayName,FirstName,LastName,ObjectId,LiveId,@{N="Onmicrosoft_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }},@{N="TargetProxyAddress_MSOL" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }}

#$Cloud = Get-EXOMailbox -ResultSize unlimited -PropertySets all

$Cloud = Get-EXOMailbox -ResultSize unlimited -Properties PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned

$cloud_filter = $Cloud  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

#$EXORecipients = Get-EXORecipient -ResultSize unlimited -PropertySets all

$EXORecipients = Get-EXORecipient -ResultSize unlimited -Properties PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,ExternalEmailAddress,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned

$Recipient_filter = $EXORecipients  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,ExternalEmailAddress,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$Mailusers = Get-MailUser -ResultSize unlimited

$Mailuser_filter = $Mailusers  | select PrimarySmtpAddress,UserPrincipalName,Alias,ExternalDirectoryObjectId,ExternalEmailAddress,RemoteRecipientType,RecipientTypeDetails,RecipientType,Name,MicrosoftOnlineServicesID,Guid,ExchangeGuid,ExchangeObjectId,DistinguishedName,EmailAddresses,CustomAttribute4,AccountDisabled,SKUAssigned,@{N="TargetProxyAddress_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="Onmicrosoft_EXO" ;E={ ($_.emailaddresses | where { $_ -match "$onmicrosoft$" }) -replace "smtp:" }}

$cloud_filter.count
