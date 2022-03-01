$MBX = "department.city@domain.de"

$user = "first.name@domain.de"

get-mailbox $MBX

get-mailbox $user

$P = Get-MailboxPermission -Identity $MBX -User $user

Remove-MailboxPermission -Identity $P.Identity -User $P.user -AccessRights $P.AccessRights -Confirm:$false -EA silentlycontinue

Add-MailboxPermission -Identity $P.Identity -User $P.user -AccessRights $P.AccessRights -AutoMapping $false

Add-RecipientPermission $P.Identity -AccessRights SendAs -Trustee $P.user -Confirm:$false -EA silentlycontinue