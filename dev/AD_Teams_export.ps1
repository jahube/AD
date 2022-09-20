
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

,@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "^AR.File" }) -split ",") -replace "CN=" }}

,@{N="Lizenzgruppe" ;E={ (($_.memberof | where { $_ -match "^AR.LIC.RR" }) -split ",")[0] -replace "CN=" }}

#Compress-Archive -Path "$REIFF_Daten" -DestinationPath "$Path\Log_Export_$ts\REIFF_Daten_($ts_short).zip" -Force # Zip Logs
#Compress-Archive -Path "$REIFF_export" -DestinationPath "$Path\Log_Export_$ts\REIFF_Telefonnummern_($ts_short).zip" -Force # Zip Logs

#Invoke-Item $Path\Log_Export_$ts # open file manager

#Unlock-ADAccount adm_windisch

$alldata = @()

Foreach ($user in $Benutzer) {

$Data = New-Object -TypeName PSObject      

$Data | Add-Member -MemberType NoteProperty -Name samaccountname -Value $user.samaccountname
#$Data | Add-Member -MemberType NoteProperty -Name msExchMailboxGuid -Value $user.msExchMailboxGuid
$Data | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $user.userprincipalname
$Data | Add-Member -MemberType NoteProperty -Name mail -Value $user.mail
$Data | Add-Member -MemberType NoteProperty -Name title -Value $user.title
#$Data | Add-Member -MemberType NoteProperty -Name employeeID -Value $user.employeeID
$Data | Add-Member -MemberType NoteProperty -Name location -Value $user.l
$Data | Add-Member -MemberType NoteProperty -Name company -Value $user.company
$Data | Add-Member -MemberType NoteProperty -Name postalcode -Value $user.postalcode
$Data | Add-Member -MemberType NoteProperty -Name streetaddress -Value $user.streetaddress
$Data | Add-Member -MemberType NoteProperty -Name telephonenumber -Value $user.telephonenumber
$Data | Add-Member -MemberType NoteProperty -Name countrycode -Value $user.countrycode
$Data | Add-Member -MemberType NoteProperty -Name department -Value $user.department
$Data | Add-Member -MemberType NoteProperty -Name description -Value $user.description
#$Data | Add-Member -MemberType NoteProperty -Name departmentnumber -Value $user.departmentnumber
#$Data | Add-Member -MemberType NoteProperty -Name Lizenzgruppe -Value $user.Lizenzgruppe
$Data | Add-Member -MemberType NoteProperty -Name OU -Value $user.OU
$Data | Add-Member -MemberType NoteProperty -Name Top_OU -Value $user.Top_OU
#$Data | Add-Member -MemberType NoteProperty -Name mailnickname -Value $user.mailnickname
#$Data | Add-Member -MemberType NoteProperty -Name msExchRemoteRecipientType -Value $user.msExchRemoteRecipientType
#$Data | Add-Member -MemberType NoteProperty -Name msExchRecipientDisplayType -Value $user.msExchRecipientDisplayType
#$Data | Add-Member -MemberType NoteProperty -Name msExchRecipientTypeDetails -Value $user.msExchRecipientTypeDetails
#$Data | Add-Member -MemberType NoteProperty -Name extensionattribute4 -Value $user.extensionattribute4
#$Data | Add-Member -MemberType NoteProperty -Name enabled -Value $user.enabled
#$Data | Add-Member -MemberType NoteProperty -Name lastlogondate -Value $user.lastlogondate
$Data | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddress -Value $user.PrimarySMTPAddress
#$Data | Add-Member -MemberType NoteProperty -Name Targetaddress -Value $user.Targetaddress
#$Data | Add-Member -MemberType NoteProperty -Name TargetProxyaddress -Value $user.TargetProxyaddress
$Data | Add-Member -MemberType NoteProperty -Name proxyaddresses -Value $user.proxyaddresses
#$Data | Add-Member -MemberType NoteProperty -Name Mailbox_Typ -Value $user.Mailbox_Typ
#$Data | Add-Member -MemberType NoteProperty -Name Serverlocation -Value $user.Serverlocation
$Data | Add-Member -MemberType NoteProperty -Name accountexpires -Value $user.accountexpires
$Data | Add-Member -MemberType NoteProperty -Name pwdLastSet -Value $user.pwdLastSet

$teams_user = $userdata | where { $_.userprincipalname -eq $user.userprincipalname }

# Teams Daten
$Data | Add-Member -MemberType NoteProperty -Name LineUri -Value $teams_user.LineUri
$Data | Add-Member -MemberType NoteProperty -Name UsageLocation -Value $teams_user.UsageLocation
$Data | Add-Member -MemberType NoteProperty -Name BPOS_S_Enterprise -Value $teams_user.BPOS_S_Enterprise
$Data | Add-Member -MemberType NoteProperty -Name BPOS_S_Standard -Value $teams_user.BPOS_S_Standard
$Data | Add-Member -MemberType NoteProperty -Name Teams -Value $teams_user.Teams
$Data | Add-Member -MemberType NoteProperty -Name MCOEV -Value $teams_user.MCOEV
$Data | Add-Member -MemberType NoteProperty -Name MCOProfessional -Value $teams_user.MCOProfessional
$Data | Add-Member -MemberType NoteProperty -Name MCO_TEAMS_IW -Value $teams_user.MCO_TEAMS_IW

# Lizenzgruppen

$User_CallingID = ($group_members | where  { $_.group -match  "^AR.Teams.CallingID" -and $_.member -eq $user.samaccountname }).group

$User_LIC = ($group_members | where  { $_.group -match  "^AR.LIC.RR.M365E3" -and $_.member -eq $user.samaccountname }).group

$User_PRT = ($group_members | where  { $_.group -match  "^AR.PRT" -and $_.member -eq $user.samaccountname }).group

$User_FILE = ($group_members | where  { $_.group -match  "^AR.FILE" -and $_.member -eq $user.samaccountname }).group

$AR_LIC_RR_M365E3_PS = $User_LIC | where  { $_ -eq "AR.LIC.RR.M365E3.PS" }

$AR_LIC_RR_M365E3_MDE = $User_LIC | where  { $_ -eq "AR.LIC.RR.M365E3.MDE" }

$AR_LIC_RR_M365E3_PSA = $User_LIC | where  { $_ -eq "AR.LIC.RR.M365E3.PSA" }

$AR_LIC_RR_M365E3_E3 = $User_LIC | where  { $_ -eq "AR.LIC.RR.M365E3.PSA" }

$E3_Lizenz = $User_LIC -join " | "

$Data | Add-Member -MemberType NoteProperty -Name E3_Lizenz -Value $E3_Lizenz
$Data | Add-Member -MemberType NoteProperty -Name M365E3_PS -Value $AR_LIC_RR_M365E3_PS
$Data | Add-Member -MemberType NoteProperty -Name M365E3_MDE -Value $AR_LIC_RR_M365E3_MDE
$Data | Add-Member -MemberType NoteProperty -Name M365E3_PSA -Value $AR_LIC_RR_M365E3_PSA
$Data | Add-Member -MemberType NoteProperty -Name M365E3_E3 -Value $AR_LIC_RR_M365E3_E3

# File Server Berechtigungen
$AR_FILE_RW = ($User_FILE | where  { $_ -match "^AR.FILE" -and  $_ -match ".RW$" }) -replace "AR.FILE." -replace ".RW" -join ","
$AR_FILE_R = ($User_FILE | where  { $_ -match "^AR.FILE" -and  $_ -match ".R$" }) -replace "AR.FILE." -replace ".R" -join ","

$Data | Add-Member -MemberType NoteProperty -Name AR_FILE_RW -Value $AR_FILE_RW
$Data | Add-Member -MemberType NoteProperty -Name AR_FILE_R -Value $AR_FILE_R

# CallingID
$AR_Teams_CallingID = $User_CallingID -replace "AR.Teams.CallingID." -join ","
$Data | Add-Member -MemberType NoteProperty -Name AR_Teams_CallingID -Value $AR_Teams_CallingID

# Printer
$AR_PRT = $User_PRT -replace "AR.PRT." -join ","
$Data | Add-Member -MemberType NoteProperty -Name AR_PRT -Value $AR_PRT

$alldata += $data

}

$TS = (get-date -Format yyyy-MM-dd_HH.mm).ToString()

$alldata | export-csv "C:\temp\AD_Export\AD_Userdaten_Komplett_Teams_Gruppen_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$alldata | select userprincipalname,samaccountname,department,description,location,telephonenumber,lineuri,Top_OU,UsageLocation,MCOEV,E3_Lizenz,M365E3_PS,M365E3_MDE,AR_Teams_CallingID,AR_FILE_RW,AR_PRT | export-csv "C:\temp\AD_Export\AD_Userdaten_SHORT_Teams_Gruppen_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation


## read teamsdata


#Update-Module MicrosoftTeams -AllowPrerelease
#Connect-MicrosoftTeams

$Teamsdata = Get-CsOnlineUser -ResultSize 2000

$userdata = $Teamsdata | select DisplayName,UserPrincipalName,LineURI,usagelocation,  @{ name = "BPOS_S_Enterprise" ; expression = { $_.AssignedPlan | where { $_.Capability -eq "BPOS_S_Enterprise" } }  },@{ name = "BPOS_S_Standard" ; expression = { $_.AssignedPlan | where { $_.Capability -eq "BPOS_S_Standard" } }  },@{ name = "Teams" ; expression = { $_.AssignedPlan | where { $_.Capability -eq "Teams" } }  },@{ name = "MCOEV" ; expression = { $_.AssignedPlan | where { $_.Capability -match "MCOEV" } }  },@{ name = "MCOProfessional" ; expression = { $_.AssignedPlan | where { $_.Capability -eq "MCOProfessional" } }  },@{ name = "MCO_TEAMS_IW" ; expression = { $_.AssignedPlan | where { $_.Capability -eq "MCO_TEAMS_IW" } }  }

## read groups

$allgroups = Get-ADGroup -Filter *

$Groups_CallingID = $allgroups | where  { $_.name -match  "^AR.Teams.CallingID" }  | select name

$Groups_LIC = $allgroups | where  { $_.name -match  "^AR.LIC.RR.M365E3" }  | select name

$Groups_PRT = $allgroups | where  { $_.name -match  "^AR.PRT" }  | select name

$Groups_FILE = $allgroups | where  { $_.name -match  "^AR.FILE" }  | select name

$Groups = @()

$Groups += $Groups_CallingID

$Groups += $Groups_LIC

$Groups += $Groups_PRT

$Groups += $Groups_FILE

$allgroups | where  { $_.name -match  "^AR.PRT" } | ft name

$allgroups | where  { $_.name -match  "^AR.FILE" } | ft name

Get-ADGroupMember AR.FILE.NL495.RW | ft SamAccountName

$group_members = @()

Foreach ($grp in $Groups) {

$memberships = Get-ADGroupMember $grp.name

Foreach ($member in $memberships) {

$item = New-Object -TypeName PSObject     
$item | Add-Member -MemberType NoteProperty -Name group -Value $grp.name
$item | Add-Member -MemberType NoteProperty -Name member -Value $member.samaccountname

$group_members += $item

}
}


