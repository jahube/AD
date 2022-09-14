
# Weiterleitungen
$Bezirk = "Nagold"

# Funktion
Get-ADUser -Filter * -Properties mail,altRecipient,deliverAndRedirect -SearchBase "$("OU=$Bezirk,OU=ELKW-Funktion,DC=elkw,DC=local")" | select mail,altRecipient,deliverAndRedirect,@{n="Target_mail" ; E= { (Get-ADObject $_.altRecipient -properties mail).mail }},@{n="Target_CN" ; E= { (Get-ADObject $_.altRecipient -Properties CN).CN }},@{n="Target_DisplayName" ; E= { (Get-ADObject $_.altRecipient -Properties DisplayName).DisplayName }},@{n="Target_Created" ; E= { (Get-ADObject $_.altRecipient -Properties Created).Created }} | Export-Csv "C:\install\$($Bezirk + "_Funktion_Weiterleitung.csv")" -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Force

# Benutzer
get-aduser -filter { mail -eq "Barbara.Dietelbach@elkw.de" } -Properties mail,altRecipient,deliverAndRedirect | select mail,altRecipient,deliverAndRedirect,@{n="Target_mail" ; E= { (Get-ADObject $_.altRecipient -properties mail).mail }},@{n="Target_CN" ; E= { (Get-ADObject $_.altRecipient -Properties CN).CN }},@{n="Target_DisplayName" ; E= { (Get-ADObject $_.altRecipient -Properties DisplayName).DisplayName }},@{n="Target_Created" ; E= { (Get-ADObject $_.altRecipient -Properties Created).Created }} 


$objectGuid  = (Get-ADUser christel.duerr).objectGuid.guid

[Convert]::ToBase64String([guid]::New("$objectGuid").ToByteArray())