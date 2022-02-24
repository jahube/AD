
Import-Module Activedirectory

#Import-Module Activedirectory

$routingdomain = "$company.mail.onmicrosoft.com"
$OU = "OU=$company-Benutzer,DC=$company,DC=local"
$prefix = "$company-Benutzer"

$MB_by_OU = Get-ADUser -SearchBase $OU -LDAPFilter "(msExchMailboxGuid=*)" -Properties msExchMailboxGuid,userprincipalname,proxyaddresses,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate

$MBXs =$MB_by_OU | select samaccountname,@{N="msExchMailboxGuid" ;E={ [GUID]($_.msExchMailboxGuid) }},userprincipalname,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,@{N="PrimarySMTPAddress" ;E={ ($_.proxyaddresses | where { $_ -cmatch "^SMTP" }) -replace "SMTP:" }},@{N="TargetAddress" ;E={ ($_.proxyaddresses | where { $_ -match "$routingdomain$" }) -replace "smtp:" }},@{N="proxyaddresses" ;E={ $_.proxyaddresses -split "," -join '|' -creplace "smtp:"  }}
$MBXs_domains = $MBXs  | select samaccountname,msExchMailboxGuid,userprincipalname,mail,msExchRemoteRecipientType,msExchRecipientDisplayType,msExchRecipientTypeDetails,extensionattribute4,enabled,lastlogondate,PrimarySMTPAddress,TargetAddress,proxyaddresses,@{N="SMTP_Domain" ;E={ ("$($_.PrimarySMTPAddress)" -split '@')[1] }},@{N="UPN_Domain" ;E={ ("$($_.userprincipalname)" -split '@')[1] }},@{N="MAIL_Domain" ;E={ ("$($_.Mail)" -split '@')[1] }},@{N="SMPT_Alias" ;E={ ($_.PrimarySMTPAddress -split '@')[0] }},@{N="UPN_Alias" ;E={ ($_.userprincipalname -split '@')[0] }}

$AccDomains = @('$company.de', 'xxx1.$company.de', 'xxx2.$company.de')

$UPN_LOCAL = $MBXs_domains | where { $_.UPN_Domain -eq "$company.local" }
$SMTP_LOCAL = $MBXs_domains | where { $_.SMTP_Domain -eq "$company.local" }

$MISSING_TARGET = $MBXs_domains | where { !($_.TargetAddress) -or ($_.TargetAddress -notmatch '@') }

$INVALID_SMTP = $MBXs_domains | where { $_.SMTP_Domain -notmatch "elkw.de$" }

$INVALID_UPN = $MBXs_domains | where { $valid_domains -notcontains $_.UPN_Domain }

$INVALID_MAIL = $MBXs_domains | where { $valid_domains -notcontains $_.MAIL_Domain }

$path = "C:\Temp"

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$L_filepath = $path + '\' + $prefix + '_ELKW_LOCAL_' + $datestamp + '.CSV'
$T_filepath = $path + '\' + $prefix + '_MISSING_TARGET_' + $datestamp + '.CSV'
$IS_filepath = $path + '\' + $prefix + '_INVALID_SMTP_' + $datestamp + '.CSV'
$IU_filepath = $path + '\' + $prefix + '_INVALID_UPN_' + $datestamp + '.CSV'
$IM_filepath = $path + '\' + $prefix + '_INVALID_Mail_' + $datestamp + '.CSV'
$M_filepath = $path + '\' + $prefix + '_MBXs_domains_' + $datestamp + '.CSV'

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