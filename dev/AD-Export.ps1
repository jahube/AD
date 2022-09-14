
$path = "c:\temp"
$Path = ([Environment]::GetFolderPath('Desktop'))

$ts = Get-Date -Format yyyy.MM.dd_HH.MM_ss
$ts_short = Get-Date -Format yyyy.MM.dd

$ELKW_Daten = mkdir "$Path\Log_Export_$ts\ELKW_Daten_($ts_short)"
$ELKW_export = mkdir "$Path\Log_Export_$ts\ELKW_export_($ts_short)"

$Benutzer_OU = "OU=ELKW-Benutzer,DC=elkw,DC=local"
$routingdomain = "elkw.mail.onmicrosoft.com"

$Benutzer_OU_liste = Get-ADOrganizationalUnit -SearchBase $Benutzer_OU -SearchScope onelevel -Filter * |  Select-Object DistinguishedName, Name

foreach ($OU in $Benutzer_OU_liste) {

#$OU_path = mkdir "$fpath\$($OU.name)"

$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -Filter * -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires
#$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

$MBXs =$MB_by_OU | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,@{n="CLIENT_ID";e= { "$([String](($_.mail -split "\.")[1] -replace "@elkw",".$(($_.mail -split "\.")[0])@ELKW").ToUpper())" }},title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},TargetAddress,@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}

$MBXs | Export-CSV -Path "$ELKW_Daten\$($OU.name)_Benutzer_detail.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$MBXs | select mail,CLIENT_ID,lastlogondate,extensionattribute4,Lizenzgruppe,enabled,Mailbox_Typ,OU,Top_OU | Export-CSV -Path "$ELKW_export\$($OU.name)_Benutzer.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
#$MBXs | Export-CSV -Path "$ELKW_Daten\$($OU.name)_Benutzer_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
#$MBXs | select mail,lastlogondate,extensionattribute4,enabled | Export-CSV -Path "$ELKW_export\$($OU.name)_Benutzer_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

}

$Funktion_OU = "OU=ELKW-Funktion,DC=elkw,DC=local"

$Funktion_OU_liste = Get-ADOrganizationalUnit -SearchBase $Funktion_OU -SearchScope onelevel -Filter * |  Select-Object DistinguishedName, Name

foreach ($OU in $Funktion_OU_liste) {

#$OU_path = mkdir "$fpath\$($OU.name)"

$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -SearchScope subtree -Filter * -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

#$MB_by_OU = Get-ADUser -SearchBase $OU.DistinguishedName -SearchScope subtree -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

$MBXs =$MB_by_OU | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="TargetAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}

$MBXs | Export-CSV -Path "$ELKW_Daten\$($OU.name)_Funktion_detail.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$MBXs | select mail,lastlogondate,extensionattribute4,enabled,Mailbox_Typ,OU,Top_OU | Export-CSV -Path "$ELKW_export\$($OU.name)_Funktion.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

#$MBXs | Export-CSV -Path "$ELKW_Daten\$($OU.name)_Funktion_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
#$MBXs | select mail,lastlogondate,extensionattribute4,enabled | Export-CSV -Path "$ELKW_export\$($OU.name)_Funktion_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

}
<#
Compress-Archive -Path "$ELKW_Daten" -DestinationPath "$Path\Log_Export_$ts\ELKW_Daten_($ts_short).zip" -Force # Zip Logs
Compress-Archive -Path "$ELKW_export" -DestinationPath "$Path\Log_Export_$ts\ELKW_export_($ts_short).zip" -Force # Zip Logs
Invoke-Item $Path\Log_Export_$ts # open file manager
#>

$ALLE_Benutzer = Get-ADUser -SearchBase $Benutzer_OU -Filter * -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

$ALLE_Benutzer.count

$ALLE_ONPREM_Benutzer = Get-ADUser -SearchBase $Benutzer_OU -LDAPFilter "(!(targetAddress=*))" -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

$ALLE_ONPREM_Benutzer.count

$ALLE_CLOUD_Benutzer = Get-ADUser -SearchBase $Benutzer_OU -LDAPFilter "(targetAddress=*)" -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,pwdlastset,accountexpires

$ALLE_CLOUD_Benutzer.count

$ONPREM_Benutzer = $ALLE_ONPREM_Benutzer| select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{N="Serverlocation" ;E={ [string]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem"};"4"{"Onprem"};"2147483648"{"Cloud"};"34359738368"{"Cloud"}}) }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}
$ONPREM_Benutzer | Export-CSV -Path "$ELKW_Daten\ALLE_ONPREM_Benutzer_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$ONPREM_Benutzer | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$ELKW_export\Alle_ONPREM_Benutzer_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$CLOUD_Benutzer = $ALLE_CLOUD_Benutzer | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{N="Serverlocation" ;E={ [string]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem"};"4"{"Onprem"};"2147483648"{"Cloud"};"34359738368"{"Cloud"}}) }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}
$CLOUD_Benutzer | Export-CSV -Path "$ELKW_Daten\ALLE_CLOUD_Benutzer_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$CLOUD_Benutzer | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$ELKW_export\Alle_CLOUD_Benutzer_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$Benutzer = $ALLE_Benutzer | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{n="CLIENT_ID";e= { "$([String](($_.mail -split "\.")[1] -replace "@elkw",".$(($_.mail -split "\.")[0])@ELKW").ToUpper())" }},@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }},@{N="Serverlocation" ;E={ [string]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem"};"4"{"Onprem"};"2147483648"{"Cloud"};"34359738368"{"Cloud"}}) }},@{name ="accountexpires";expression={[datetime]::FromFileTime($_.accountexpires)}},@{name ="pwdLastSet";expression={[datetime]::FromFileTime($_.pwdLastSet)}}
$Benutzer | Export-CSV -Path "$ELKW_Daten\ALLE_BENUTZER_Daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Benutzer | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,CLIENT_ID,extensionattribute4,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$ELKW_export\ALLE_BENUTZER_Export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Benutzer | select samaccountname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,lastlogondate,CLIENT_ID,extensionattribute4,Lizenzgruppe,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$ELKW_export\ALLE_BENUTZER_VALUEMATION_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$ALLE_Pfarrer = $ALLE_Benutzer | where { $_.title -match "Pfarrer"}
$Pfarrer = $ALLE_Pfarrer | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},@{N="Top_OU" ;E={ [String]((($_.DistinguishedName) -split ",")[-4] -replace "OU=")}},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ [String]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}}) }},@{N="Serverlocation" ;E={ [string]$(Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem"};"4"{"Onprem"};"2147483648"{"Cloud"};"34359738368"{"Cloud"}}) }}
$Pfarrer | Export-CSV -Path "$ELKW_Daten\Alle_Pfarrer_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Pfarrer | select samaccountname,mail,title,employeeID,lastlogondate,extensionattribute4,enabled,Serverlocation,OU,Top_OU | Export-CSV -Path "$ELKW_export\ALLE_Pfarrer_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$ALLE_ONPREM_Funktionspostfaecher = Get-ADUser -SearchBase $Funktion_OU -LDAPFilter "(!(targetAddress=*))" -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$ALLE_ONPREM_Funktionspostfaecher.count

($ALLE_ONPREM_Funktionspostfaecher |where { $_.enabled -eq $true }).count

$ALLE_CLOUD_Funktionspostfaecher = Get-ADUser -SearchBase $Funktion_OU -LDAPFilter "(targetAddress=*)" -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$ALLE_CLOUD_Funktionspostfaecher.count

$ALLE_Funktionspostfaecher = Get-ADUser -SearchBase $Funktion_OU -Filter * -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$ALLE_Funktionspostfaecher.count

$ONPREM_Funktionspostfaecher = $ALLE_ONPREM_Funktionspostfaecher| select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ [String](($_.proxyaddresses | where { $_ -match "$routingdomain" }) -replace "smtp:") }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }}
$ONPREM_Funktionspostfaecher | Export-CSV -Path "$ELKW_Daten\ALLE_ONPREM_Funktionspostfaecher_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$ONPREM_Funktionspostfaecher | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Mailbox_Typ,OU,Top_OU | Export-CSV -Path "$ELKW_export\Alle_ONPREM_Funktionspostfaecher_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$CLOUD_Funktionspostfaecher = $ALLE_CLOUD_Funktionspostfaecher| select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ [String](($_.proxyaddresses | where { $_ -match "$routingdomain" }) -replace "smtp:") }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }}
$CLOUD_Funktionspostfaecher | Export-CSV -Path "$ELKW_Daten\ALLE_CLOUD_Funktionspostfaecher_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$CLOUD_Funktionspostfaecher | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Mailbox_Typ,OU,Top_OU | Export-CSV -Path "$ELKW_export\Alle_CLOUD_Funktionspostfaecher_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$Funktionspostfaecher = $ALLE_Funktionspostfaecher | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }}
$Funktionspostfaecher | Export-CSV -Path "$ELKW_Daten\Alle_Funktionspostfaecher_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$Funktionspostfaecher | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Mailbox_Typ,OU,Top_OU | Export-CSV -Path "$ELKW_export\ALLE_Funktionspostfaecher_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation


$Group_List =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3","ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E3_Sekretaerinnen", "ELKW_O365_LIC_M3_Synode"

$MemberlistLizenzen = foreach ($licgroup in $Group_List) { Get-ADGroupMember $licgroup }

[System.Collections.ArrayList]$LizenzGuppenuser = $MemberlistLizenzen.SamAccountname

$LizenzGuppenuser.count

$LizenzGuppenuser_array = foreach ($Licuser in  $LizenzGuppenuser) {

Get-ADUser $Licuser -Properties msExchMailboxGuid,userprincipalname,mailnickname,proxyaddresses,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,departmentnumber,memberof,TargetAddress,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

}

$LizenzGuppen_user = $LizenzGuppenuser_array | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,title,employeeID,l,company,postalcode,streetaddress,telephonenumber,countrycode,@{N="departmentnumber" ;E={ [String]($_.departmentnumber) }},@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "O365" }) -split ",")[0] -replace "CN=" }},@{N="OU" ;E={ (($_.DistinguishedName) -split "," -replace "OU=")[(($_.DistinguishedName -split ",").count - 3)..2] -join "/" }},mailnickname,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="Targetaddress" ;E={ [String](($_.Targetaddress ) -replace "smtp:") }},@{N="TargetProxyaddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }},@{N="Mailbox_Typ" ;E={ Switch ($_.msExchRecipientTypeDetails){"1"{"Onprem_Benutzer"};"4"{"Onprem_Funktion"};"2147483648"{"Cloud_Benutzer"};"34359738368"{"Cloud_Funktion"}} }}
$LizenzGuppen_user | Export-CSV -Path "$ELKW_Daten\Alle_LizenzGuppenuser_daten_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$LizenzGuppen_user | select samaccountname,mail,title,employeeID,Lizenzgruppe,lastlogondate,extensionattribute4,enabled,Mailbox_Typ,OU | Export-CSV -Path "$ELKW_export\ALLE_LizenzGuppenuser_export_($ts_short).CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation
$LizenzGuppen_user.count

Compress-Archive -Path "$ELKW_Daten" -DestinationPath "$Path\Log_Export_$ts\ELKW_Daten_($ts_short).zip" -Force # Zip Logs
Compress-Archive -Path "$ELKW_export" -DestinationPath "$Path\Log_Export_$ts\ELKW_export_($ts_short).zip" -Force # Zip Logs
Invoke-Item $Path\Log_Export_$ts # open file manager

# routing mismatch proxy / targetaddress - Benutzer
$rout_error_Benutzer = $CLOUD_Benutzer | where { $_.TargetAddress -ne $_.TargetProxyaddress }

$rout_error_Benutzer.count

$rout_error_Benutzer  |select samaccountname,mail,title,employeeID,memberof,TargetAddress,TargetProxyaddress,OU,@{N="removeproxy" ;E={ [String](($_.TargetProxyaddress -replace $_.TargetAddress -replace " ").Trim()) }}

# routing mismatch proxy / targetaddress - Funktion
$rout_error_funktion = $CLOUD_Funktionspostfaecher | where { $_.TargetAddress -ne $_.TargetProxyaddress }

$rout_error_funktion.count

$rout_error_funktion  |select samaccountname,mail,title,employeeID,memberof,TargetAddress,TargetProxyaddress,OU,@{N="removeproxy" ;E={ [String](($_.TargetProxyaddress -replace $_.TargetAddress -replace " ").Trim()) }}

# mismatch Mail / PrimarySMTP  - Benutzer
$sMTP_mismatch_Benutzer = $Benutzer | where { $_.mail -ne $_.PrimarySMTPAddress }

$sMTP_mismatch_Benutzer.count

$sMTP_mismatch_Benutzer| ft mail,PrimarySMTPAddress,OU,Mailbox_Typ -AutoSize

$sMTP_mismatch_Benutzer| select samaccountname,mail,title,employeeID,memberof,PrimarySMTPAddress,OU,Mailbox_Typ

# mismatch Mail / PrimarySMTP  - Funktion

$sMTP_mismatch_Funktion = $Funktionspostfaecher | where  { $_.mail -ne $_.PrimarySMTPAddress }

$sMTP_mismatch_Funktion.count

$sMTP_mismatch_Funktion| ft mail,PrimarySMTPAddress,OU,Mailbox_Typ -AutoSize

$sMTP_mismatch_Funktion| select samaccountname,mail,title,employeeID,memberof,PrimarySMTPAddress,OU,Mailbox_Typ

<#
######################################################
# Fix Benutzer
######################################################

[System.Collections.ArrayList]$fix_Benutzer =   $rout_error_Benutzer |select samaccountname,mail,title,employeeID,memberof,TargetAddress,TargetProxyaddress,@{N="removeproxy" ;E={ [String](($_.TargetProxyaddress -replace $_.TargetAddress -replace " ").Trim()) }}

$fix_Benutzer | ft samaccountname,TargetAddress,removeproxy -AutoSize

$fix_Benutzer = [System.Collections.ArrayList]$fix_Benutzer

$fix_Benutzer.RemoveAt([array]::indexof($fix.samaccountname,'USER.NAME'))

$fix_Benutzer | where { $_.TargetAddress -eq $_.removeproxy } | ft samaccountname,TargetAddress,removeproxy -AutoSize


foreach ($user in $fix_Benutzer) {

$removeaddress = 'smtp:'+ ($User.removeproxy).Trim()
$removeaddress
#$User.removeproxy
#set-aduser $user.samaccountname -remove @{ProxyAddresses= $removeaddress}
}

######################################################
#   Fix Funktionspostfächer
######################################################

[System.Collections.ArrayList]$fix_Benutzer =   $rout_error_funktion |select samaccountname,mail,title,employeeID,memberof,TargetAddress,TargetProxyaddress,@{N="removeproxy" ;E={ [String](($_.TargetProxyaddress -replace $_.TargetAddress -replace " ").Trim()) }}

$fix_Funktionspostfaecher | ft samaccountname,TargetAddress,removeproxy -AutoSize

$fix_Funktionspostfaecher = [System.Collections.ArrayList]$fix_Funktionspostfaecher

$fix_Funktionspostfaecher.RemoveAt([array]::indexof($Funktionspostfaecher.samaccountname,'USER.NAME'))

$fix_Funktionspostfaecher | where { $_.TargetAddress -eq $_.removeproxy } | ft samaccountname,TargetAddress,removeproxy -AutoSize


foreach ($user in $fix_Funktionspostfaecher) {

$removeaddress = 'smtp:'+ ($User.removeproxy).Trim()
$removeaddress
#$User.removeproxy
#set-aduser $user.samaccountname -remove @{ProxyAddresses= $removeaddress}
}

#>