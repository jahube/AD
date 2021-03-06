$OU = "OU=ccc-Benutzer,DC=ccc,DC=local"
$Group_List =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3_Default", "ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E3_Sekretaerinnen"

#$MB_by_OU = Get-ADUser -SearchBase $OU -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,mailnickname,TargetAddress,ObjectGUID,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$MB_by_OU = Get-ADUser -Filter * -Properties memberof,msExchMailboxGuid,userprincipalname,proxyaddresses,mail,mailnickname,TargetAddress,ObjectGUID,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,Distinguishedname

$MBXs1 = $MB_by_OU.where({ (($_.Distinguishedname.Split(','))[-3]) -match "Benutzer$" -or (($_.Distinguishedname.Split(','))[-3])  -match "Funktion$" })

$MBXs = $MBXs1 | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,mailnickname,TargetAddress,ObjectGUID,@{N="Lizenzgruppen" ;E={ (((($_.memberof.Split(','))[0]) -replace "CN=") | where { $_ -in $Group_List }) -join '|' }},msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ (($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:") }},@{N="TargetProxyAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ ($_.proxyaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:" }},@{N="proxyaddresses_join" ;E={ (($_.proxyaddresses | where { $_ -notmatch "^x500" -and $_ -notmatch "^SIP" -and $_ -notmatch "^SPO" }) -creplace "smtp:") -join '|' }},@{N="Top-Level_OU" ;E={ (($_.Distinguishedname.Split(','))[-4]) -replace "OU=" }},Distinguishedname,@{N="SMTP_Domain" ;E={ ((($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:") -split '@')[1] }},@{N="UPN_Domain" ;E={ ("$($_.userprincipalname)" -split '@')[1] }},@{N="MAIL_Domain" ;E={ ("$($_.Mail)" -split '@')[1] }},@{N="SMTP_Alias" ;E={ ((($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:") -split '@')[0] }},@{N="UPN_Alias" ;E={ ($_.userprincipalname -split '@')[0] }},deliverAndRedirect,@{n="altRecipient_mail" ; E= { (Get-ADObject $_.altRecipient -properties mail).mail }}

$AccDomains = @('x1.de', 'x2.de')

$UPN_LOCAL = $MBXs_domains | where { $_.UPN_Domain -eq "elkw.local" }
$SMTP_LOCAL = $MBXs_domains | where { $_.SMTP_Domain -eq "elkw.local" }

$MISSING_TARGET = $MBXs_domains | where { !($_.TargetAddress) -or ($_.TargetAddress -notmatch '@') }

$INVALID_SMTP = $MBXs_domains | where { $_.SMTP_Domain -notmatch "elkw.de$" }

$INVALID_UPN = $MBXs_domains | where { $valid_domains -notcontains $_.UPN_Domain }

$INVALID_MAIL = $MBXs_domains | where { $valid_domains -notcontains $_.MAIL_Domain }

$path = "C:\Temp"

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$L_filepath = $path + '\' + $prefix + '_ELKW_LOCAL_' + $($Bereich) + '_' + $datestamp + '.CSV'
$T_filepath = $path + '\' + $prefix + '_MISSING_TARGET_' + $($Bereich) + '_' + $datestamp + '.CSV'
$IS_filepath = $path + '\' + $prefix + '_INVALID_SMTP_' + $($Bereich) + '_' + $datestamp + '.CSV'
$IU_filepath = $path + '\' + $prefix + '_INVALID_UPN_' + $($Bereich) + '_' + $datestamp + '.CSV'
$IM_filepath = $path + '\' + $prefix + '_INVALID_Mail_' + $($Bereich) + '_' + $datestamp + '.CSV'
$M_filepath = $path + '\' + $prefix + '_MBXs_domains_' + $($Bereich) + '_' + $datestamp + '.CSV'

$ELKW_LOCAL | Export-Csv -Path $L_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
$MISSING_TARGET | Export-Csv -Path $T_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
$INVALID_SMTP | Export-Csv -Path $IS_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
$INVALID_UPN  | Export-Csv -Path $IU_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
$INVALID_MAIL | Export-Csv -Path $IM_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
$MBXs_domains | Export-Csv -Path $M_filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force

$ELKW_LOCAL.count
$MISSING_TARGET.count
$INVALID_SMTP.count
$INVALID_UPN.count
$INVALID_MAIL.count
$MBXs_domains.count

$INVALID_SMTP | ft user*,prim*
$INVALID_UPN | ft user*,prim*
$INVALID_MAIL | ft user*,prim*


$upn_fail_count = $MBXs_domains | group upn_domain | select name,count | sort-object count -Descending
$smtp_fail_count = $MBXs_domains | group smtp_domain | select name,count | sort-object count -Descending
$mail_fail_count = $MBXs_domains | group mail_domain | select name,count | sort-object count -Descending

$upn_fail_count.count
$smtp_fail_count.count
$mail_fail_count.count


$upn_fail_count
$smtp_fail_count
$mail_fail_count