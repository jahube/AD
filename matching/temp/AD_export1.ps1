
$OU = "OU=ELKW-Benutzer,DC=elkw,DC=local"

$MB_by_OU = Get-ADUser -SearchBase $OU -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,mailnickname,TargetAddress,ObjectGUID,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$MBXs =$Benutzer_filter | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,mailnickname,TargetAddress,ObjectGUID,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="TargetProxyAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},proxyaddresses,@{N="proxyaddresses_joined" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }}
