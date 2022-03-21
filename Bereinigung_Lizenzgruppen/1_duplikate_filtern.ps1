$Group_List1 =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3", "ELKW_O365_LIC_E3_Sekretaerinnen"

$MemberlistLizenzen = foreach ($licgroup in $Group_List1) { Get-ADGroupMember $licgroup }

[System.Collections.ArrayList]$samaccountnames = $MemberlistLizenzen.SamAccountName | sort-object -unique

$mbx = foreach ($mbx in $samaccountnames) { get-aduser $mbx -Properties mail,extensionattribute4,memberof,Distinguishedname }

$duplikate = $mbx.where({$_.extensionattribute4 -match "^AP-Typ"})

$duplikate.count

$Group_List2 =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3","ELKW_O365_LIC_E3_Default","ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E3_Sekretaerinnen"

$data = $duplikate | select mail,extensionattribute4,samaccountname,@{N="Top-Level_OU" ;E={ (($_.Distinguishedname.Split(','))[-4]) -replace "OU=" }},@{N="Lizenzgruppen" ;E={ (Get-ADPrincipalGroupMembership $_.SamAccountName |where { $_.Name -in $Group_List2 }).SamAccountName -join '|' }},@{N="memberof" ;E={ (($_.memberof.Split(','))[0]) -replace "cn=" -join ' | ' }},Distinguishedname

$data |ft mail,extensionattribute4,samaccountname,Top-Level_OU,Lizenzgruppen,memberof

foreach ($user in $data) {

foreach ($Group in ($User.Lizenzgruppen -Split '|')) {

Remove-ADGroupMember -Identity $User.Lizenzgruppen -Members $User.SamAccountName -Confirm:$false 
                                                     }
                         }

## end

#foreach ($user in $data) { Remove-ADGroupMember -Identity ELKW_O365_LIC_E3 -Members $User.SamAccountName -Confirm:$false  }