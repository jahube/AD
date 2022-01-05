########### Anpassen ################################################################################

Get-mailbox -filter { userprincipalname -like "*NAME*" } | select userprincipalname,primarysmtpaddress

$u = get-mailbox "User1_ONPREM@domain.de"

$m = get-mailbox "FunktionONPREM@domain.de"

#####################################################################################################

# list Permission

$m | Get-MailboxPermission | where {$_.isinherited -eq $false }

$m | Get-ADPermission | where {$_.isinherited -eq $false -and $_.User -notlike "NT-Aut*"}

#####################################################################################################

Add-mailboxpermission $m.distinguishedname -user $u.distinguishedname -Accessrights FullAccess

# Automapping 
Set-ADUser -Identity $m.distinguishedname -Add @{msExchDelegateListLink="$($u.distinguishedname)"}

# Backlink im user ZUM geteilten Postfach prüfen
Get-ADUser -Identity $u.distinguishedname -properties msExchDelegateListBL | select -ExpandProperty msExchDelegateListBL

# SendAs         ("Senden Als")
Add-ADPermission -Identity $m.distinguishedname -user $u.distinguishedname -ExtendedRights "Send As"

# Send on Behalf ("Senden im Namen von")
set-mailbox -Identity $m.distinguishedname -grantsendonbehalfto @{Add="$($u.distinguishedname)"}

#####################################################################################################
# nur für Korrekturen
#####################################################################################################

<# clear automapping
## Funktionspostfach

# beim Funktionspostfach Automapping für obigen USER rausnehmen
Set-ADUser -Identity $m.distinguishedname -Remove @{msExchDelegateListLink="$($u.distinguishedname)"}

# beim Funktionspostfach Automapping für ALLE rausnehmen
Set-ADUser -Identity $m.distinguishedname -clear msExchDelegateListLink


## USER
# beim USER Automapping für rausnehmen (eigentlich "falsche" Richtung)
Set-ADUser -Identity $u.distinguishedname -clear msExchDelegateListLink
#>
#####################################################################################################