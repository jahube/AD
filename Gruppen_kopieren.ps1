
$Vorlage_user = "Vorname.name"

$user1 = "Vorname.name"

$gruppen_vorlage = (get-aduser $Vorlage_user -Properties memberof).memberof

$gruppen_User1= (get-aduser $user1 -Properties memberof).memberof

$gruppen_vorlage.count

$gruppen_User1.count

foreach ($gruppe in $gruppen) {

If(!($gruppe -in $gruppen_User1)) {

Add-ADGroupMember $gruppe -Members $user1

} }

# check
$gruppen_User1= (get-aduser $user1 -Properties memberof).memberof
$gruppen_User1.count