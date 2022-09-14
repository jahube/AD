
$Group_List =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3","ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E3_Sekretaerinnen","ELKW_O365_LIC_M3_Synode"

$MemberlistLizenzen = foreach ($licgroup in $Group_List) { Get-ADGroupMember $licgroup }
$MemberlistLizenzen.count
$mbx = get-aduser -F * -Properties mail,extensionattribute4,memberof

[System.Collections.ArrayList]$cloudlizenzen = ($mbx.where({$_.extensionattribute4 -match "^AP-Typ"})).SamAccountname

$cloudlizenzen.count

[System.Collections.ArrayList]$LizenzGuppenuser = $MemberlistLizenzen.SamAccountname

$LizenzGuppenuser.count

$duplikate = $LizenzGuppenuser | where {$_ -in $cloudlizenzen }

$duplikate.count

$data = $duplikate | % { get-aduser $_ -properties mail,extensionattribute4,samaccountname,memberof,Distinguishedname }

$data | select mail,extensionattribute4,samaccountname,@{N="Top-Level_OU" ;E={ (($_.Distinguishedname.Split(','))[-4]) -replace "OU=" }},`
@{N="memberof" ;E={ (($_.memberof.Split(','))[0]) -replace "cn=" -join ' | ' }} |ft -AutoSize

############################################################################

$Users = $data.mail

$ts = get-date -Format yyyyMMdd_HHmm_ss

$LogFile = "C:\Users\adm_huebener\Documents\t\Logs_$ts.txt"

$LicGroups = (Get-Adgroup -Filter { Name -like "*O365*"} -searchbase "OU=Datagroup,OU=ELKW-Benutzer,DC=elkw,DC=local").Name

$Group_List =  "ELKW_O365_LIC_M3_Default","ELKW_O365_LIC_E3_Default", "ELKW_O365_LIC_E1_Default", "ELKW_O365_LIC_E3_Sekretaerinnen","ELKW_O365_LIC_M3_Synode"
#$Group_List = (Get-Adgroup -Filter { Name -like "*O365*"} -searchbase "OU=Datagroup,OU=ELKW-Benutzer,DC=elkw,DC=local").Name

foreach ($User in $Users)
{

# Retrieve UPN
    $UPN =  $User.Trim()
    Write-Host -ForegroundColor Gray "Processing $UPN..."

    # Retrieve UPN related SamAccountName
    #$ADUser = Get-ADUser -Filter {UserPrincipalname -eq $UPN} | Select-Object SamAccountName

    $ADUser = Get-ADUser -Filter { Mail -eq $UPN} | Select-Object SamAccountName
    
    # User from CSV not in AD
    if ($ADUser -eq $null) {     
      Write-Host "$UPN does not exist in AD`n" -ForegroundColor Red
    #  Start-Sleep 2
                           } 
                           else {
        # Retrieve AD user group membership
        $License_Groups = Get-ADPrincipalGroupMembership $ADUser.SamAccountName |where { $_.Name -in $Group_List }

       # $LIC_G_found = ($ExistingGroups | where { $_.Name -match "ELKW_O365_LIC" }).Name

                           }

#IF($ExistingGroups -and $ADUser){     #$LIC_G_found

       if (!($License_Groups)) {
            # User not member of group
            Write-Host "$UPN does not exist in any Lic Group!" -ForeGroundColor Yellow

            $Time = Get-Date -Format "[hh:mm:ss]:"
            Write-Output "$Time User aus keiner Gruppe entfernt $($user.UserPrincipalName)" >> $LogFile
        }

   foreach ($Group in $License_Groups)
  {
        
        try { 

        Remove-ADGroupMember -Identity $Group.SamAccountName -Members $ADUser.SamAccountName -Confirm:$false -EA stop

                  Write-Host "Removed $UPN from $($Group.name)" -ForeGroundColor Green

                Write-Output "$Time User aus Gruppe entfernt $($ADUser.SamAccountName)" >> $LogFile

            } 

        catch { 

                  write-host $Error[0].exception.message -F yellow 

                Write-Output "$Time Entfernen $UPN aus Gruppe $($ADUser.SamAccountName) entfernen FAILED" >> $LogFile 
        }

   

   }
} 
