
$path = "c:\temp"
$DesktopPath = ([Environment]::GetFolderPath('Desktop'))
$company = "xxxx"

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM_ss

$Benutzer_logs = mkdir "$path\Benutzer_export_$datestamp"

$Benutzer_OU = "OU=$company-Benutzer,DC=$company,DC=local"
#$Benutzer_OU = "OU=Test Accounts,DC=pshell,DC=site"

$Benutzer_OU_liste = Get-ADOrganizationalUnit -SearchBase $Benutzer_OU -SearchScope onelevel -Filter * |  Select-Object DistinguishedName, Name

foreach ($OU in $Benutzer_OU_liste) {

#$OU_path = mkdir "$fpath\$($OU.name)"

$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$MBXs =$MB_by_OU | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="TargetAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }}

$MBXs | Export-CSV -Path "$Benutzer_logs\$($OU.name).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$MBXs | select mail,lastlogondate,extensionattribute4 | Export-CSV -Path "$Benutzer_logs\$($OU.name)_SHORT.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation



}


$funktion_logs = mkdir "$path\Funktion_export_$datestamp"

$Funktion_OU = "OU=$company-Funktion,DC=$company,DC=local"

$Funktion_OU_liste = Get-ADOrganizationalUnit -SearchBase $Funktion_OU -SearchScope onelevel -Filter * |  Select-Object DistinguishedName, Name

foreach ($OU in $Funktion_OU_liste) {

#$OU_path = mkdir "$fpath\$($OU.name)"

$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -SearchScope subtree -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$MBXs =$MB_by_OU | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="TargetAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }}

$MBXs | Export-CSV -Path "$funktion_logs\$($OU.name).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$MBXs | select mail,lastlogondate,extensionattribute4 | Export-CSV -Path "$funktion_logs\$($OU.name)_SHORT.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

}