$mail = "USER_MAIL"

Get-ADUser -Filter {mail -eq $mail} -Properties extensionAttribute4 | select name, extensionAttribute4 | ft

Get-ADUser -Filter {mail -eq $mail } | set-ADUser -replace @{extensionAttribute4 = "AP-Type-14"}

Get-ADUser -Filter {mail -eq $mail} -Properties extensionAttribute4 | select name, extensionAttribute4 | ft