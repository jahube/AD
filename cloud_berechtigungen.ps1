$u = get-mailbox "User1_CLOUD@domain.de"

$m = get-mailuser "FunktionONPREM@domain.de"

Add-recipientPermission $u.distinguishedname -Trustee $m.distinguishedname -AccessRights SendAs