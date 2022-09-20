
$path = "C:\temp\AD_Export"

$ts = Get-Date -Format yyyy.MM.dd_HH.MM_ss
$ts_short = Get-Date -Format yyyy.MM.dd

#$REIFF_Daten = mkdir "$Path\Log_Export_$ts\REIFF_Daten_($ts_short)"
#$REIFF_Telefon = mkdir "$Path\Log_Export_$ts\REIFF_Telefonnummern_($ts_short)"

$Benutzer_OU = "OU=OU_Reiff,DC=reiff-reifen,DC=de"
$routingdomain = "reiffreifende.onmicrosoft.com"

$ALLE_Benutzer = Get-ADUser -SearchBase $Benutzer_OU -Filter * -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,department,description,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires
$ALLE_Benutzer.count

$Benutzer = $ALLE_Benutzer | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,department,description,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "^AR.LIC.RR" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{n="CLIENT_ID";e= { "$([String](($_.mail -split "\.")[1] -replace "@elkw",".$(($_.mail -split "\.")[0])@ELKW").ToUpper())" }},@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{N="Serverlocation" ;E={ [string]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem"};"4"{"Onprem"};"2147483648"{"Cloud"};"34359738368"{"Cloud"}}) }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}
$Benutzer | Export-CSV -Path "$path\ALLE_BENUTZER_Daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Benutzer | select userprincipalname,telephonenumber,department,description,samaccountname,mail,title,employeeID,lastlogondate,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$path\REIFF_USER_Details_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Benutzer | select userprincipalname,telephonenumber,department | Export-CSV -Path "$path\REIFF_USER_Telefonnummern_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

#Compress-Archive -Path "$REIFF_Daten" -DestinationPath "$Path\Log_Export_$ts\REIFF_Daten_($ts_short).zip" -Force # Zip Logs
#Compress-Archive -Path "$REIFF_export" -DestinationPath "$Path\Log_Export_$ts\REIFF_Telefonnummern_($ts_short).zip" -Force # Zip Logs

#Invoke-Item $Path\Log_Export_$ts # open file manager

#Unlock-ADAccount adm_windisch