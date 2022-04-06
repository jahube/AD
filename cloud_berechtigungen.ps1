$m = get-mailbox "FunktionCLOUD@domain.de"

$u = get-mailuser "User1_ONPREM@domain.de"

Add-recipientPermission $m.distinguishedname -Trustee $u.distinguishedname -AccessRights SendAs

Get-recipientPermission $m.distinguishedname

--------------------------------------------

$m = get-mailuser "FunktionONPREM@domain.de"

$u = get-mailbox "User1_CLOUD@domain.de"

Add-recipientPermission $m.distinguishedname -Trustee $u.distinguishedname -AccessRights SendAs

Get-recipientPermission $m.distinguishedname