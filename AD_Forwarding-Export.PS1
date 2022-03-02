# Weiterleitungen
$Bezirk = "Bezirk"

$c_cap  = "EL**"

$c_small = "el**"

$OU_Funktion = "$("OU=$Bezirk,OU=$c_cap-Funktion,DC=$c_small,DC=local")"

$OU_Benutzer = "$("OU=$Bezirk,OU=$c_cap-Benutzer,DC=$c_small,DC=local")"

# Funktion
Get-ADUser -Filter * -Properties mail,altRecipient,deliverAndRedirect -SearchBase $OU_Funktion | select mail,altRecipient,deliverAndRedirect,@{n="Target_mail" ; E= { (Get-ADObject $_.altRecipient -properties mail).mail }},@{n="Target_CN" ; E= { (Get-ADObject $_.altRecipient -Properties CN).CN }},@{n="Target_DisplayName" ; E= { (Get-ADObject $_.altRecipient -Properties DisplayName).DisplayName }},@{n="Target_Created" ; E= { (Get-ADObject $_.altRecipient -Properties Created).Created }} | Export-Csv "C:\install\$($Bezirk + "_Funktion_Weiterleitung.csv")" -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Force

# Benutzer
Get-ADUser -Filter * -Properties mail,altRecipient,deliverAndRedirect -SearchBase $OU_Benutzer | select mail,altRecipient,deliverAndRedirect,@{n="Target_mail" ; E= { (Get-ADObject $_.altRecipient -properties mail).mail }},@{n="Target_CN" ; E= { (Get-ADObject $_.altRecipient -Properties CN).CN }},@{n="Target_DisplayName" ; E= { (Get-ADObject $_.altRecipient -Properties DisplayName).DisplayName }},@{n="Target_Created" ; E= { (Get-ADObject $_.altRecipient -Properties Created).Created }} | Export-Csv "C:\install\$($Bezirk + "_Benutzer_Weiterleitung.csv")" -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Force
