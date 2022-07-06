########### Anpassen ################################################################################

Get-mailbox -filter { userprincipalname -like "*NAME*" } | select userprincipalname,primarysmtpaddress

$u = get-mailbox "User@domain.de"

$m = get-mailbox "Funktion@domain.de"

#####################################################################################################

# list existing old Permission

# Full Access
$m | Get-MailboxPermission | where {$_.isinherited -eq $false -and $_.User -notlike "NT-Aut*"}

# SendAs
$m | Get-ADPermission | where {$_.isinherited -eq $false -and $_.User -notlike "NT-Aut*"}

#####################################################################################################

Add-mailboxpermission $m.distinguishedname -user $u.distinguishedname -Accessrights FullAccess -AutoMapping $false -Confirm:$false

# Remove-mailboxpermission $m.distinguishedname -user $u.distinguishedname -Accessrights FullAccess -Confirm:$false

# Automapping 
Set-ADUser -Identity $m.distinguishedname -Add @{msExchDelegateListLink="$($u.distinguishedname)"}

# Backlink im user ZUM geteilten Postfach pruefen
Get-ADUser -Identity $u.distinguishedname -properties msExchDelegateListBL | select -ExpandProperty msExchDelegateListBL

# SendAs         ("Senden Als")
Add-ADPermission -Identity $m.distinguishedname -user $u.distinguishedname -ExtendedRights "Send As" -Confirm:$false

# SendAs CLOUD   ("Senden Als")
Add-recipientPermission $m.distinguishedname -Trustee $u.distinguishedname -AccessRights SendAs -Confirm:$false

# Send on Behalf ("Senden im Namen von")
set-mailbox -Identity $m.distinguishedname -grantsendonbehalfto @{Add="$($u.distinguishedname)"}

#####################################################################################################
# nur fuer Korrekturen
#####################################################################################################

<# clear automapping
## Funktionspostfach

# beim Funktionspostfach Automapping fuer obigen USER rausnehmen
Set-ADUser -Identity $m.distinguishedname -Remove @{msExchDelegateListLink="$($u.distinguishedname)"}

# beim Funktionspostfach Automapping für ALLE rausnehmen
Set-ADUser -Identity $m.distinguishedname -clear msExchDelegateListLink


## USER
# beim USER Automapping auf Funktionspostfach rausnehmen ("falsche" Richtung)
Set-ADUser -Identity $u.distinguishedname -clear msExchDelegateListLink
#>
#####################################################################################################