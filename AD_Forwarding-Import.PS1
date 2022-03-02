# Weiterleitungen
$Bezirk = "Bezirk"

$data = Import-Csv "C:\install\$($Bezirk + "_Funktion_Weiterleitung.csv")" -Delimiter ";" -Encoding UTF8

$forward_data = $data |where { $_.Target_mail -match "@"}

foreach ($U in $forward_data) {

Set-mailbox $U.mail -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $U.Target_mail -Confirm:$false

#New-InboxRule -Mailbox $U.mail -ForwardTo $U.Target_mail 

}