 $mbxs = get-mailbox -resultsize unlimited

                  $Data = @()

               $Orphans = @()

$count = $mbxs.count

[Int]$Num = 1

foreach ($M in $mbxs) {

$D = $M.distinguishedname ; $A = $M.alias ; $U = $M.PrimarySmtpAddress.ToString()

$OUT = get-MailboxPermission $D

$USERTYPE = $Null

$result = $OUT | where { $_.isinherited -eq $false -and "$($_.user)" -notmatch "anonym" -and "$($_.user)" -ne "default" -and  "$($_.user)" -ne "Standard" -and $_.user -notlike "NT*Auth*\SEL*" }

IF ($result) {

foreach ($P in $result ) {

IF ($P.user -like "*S-1-5-21*") {

$USERTYPE = "Unknown"

            $item = New-Object -TypeName PSCustomObject
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_UPN" -Value  $M.userprincipalname
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_SMTP" -Value  $M.PrimarySMTPAddress
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_Alias" -Value $M.Alias
            $item | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $M.SamAccountName
            $item | Add-Member -MemberType NoteProperty -Name "M_TargetAddress" -Value "$(($M.emailaddresses | where { $_ -match ".mail.onmicrosoft.com" }) -replace "smtp:")"
            $item | Add-Member -MemberType NoteProperty -Name "M_OU_DN" -Value "$((($M.DistinguishedName) -split "," -replace "OU=")[(($M.DistinguishedName -split ",").count - 3)..2] -join "/")"
            $item | Add-Member -MemberType NoteProperty -Name "M_OU" -Value $M.Organizationalunit
            $item | Add-Member -MemberType NoteProperty -Name "M_DATABASE" -Value $M.DataBase
            $item | Add-Member -MemberType NoteProperty -Name "M_Exchangeguid" -Value $M.Exchangeguid
            $item | Add-Member -MemberType NoteProperty -Name "M_Guid" -Value $M.Guid
            $item | Add-Member -MemberType NoteProperty -Name "Identity" -Value $P.Identity
            $item | Add-Member -MemberType NoteProperty -Name "USERTYPE" -Value "Unknown"
            $item | Add-Member -MemberType NoteProperty -Name "User" -Value "$($P.user)"
            $item | Add-Member -MemberType NoteProperty -Name "AccessRights" -Value "$({$P.AccessRights})"
            $item | Add-Member -MemberType NoteProperty -Name "SharingPermissionFlags" -Value "$({$P.SharingPermissionFlags})"

$Orphans += $item

} ELSE {

             $U = (get-recipient $P.user -EA silentlycontinue)

            $item = New-Object -TypeName PSCustomObject
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_UPN" -Value  $M.userprincipalname
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_SMTP" -Value  $M.PrimarySMTPAddress
            $item | Add-Member -MemberType NoteProperty -Name "Mailbox_Alias" -Value $M.Alias
            $item | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $M.SamAccountName
            $item | Add-Member -MemberType NoteProperty -Name "M_DATABASE" -Value $M.DataBase
            $item | Add-Member -MemberType NoteProperty -Name "M_Exchangeguid" -Value $M.Exchangeguid
            $item | Add-Member -MemberType NoteProperty -Name "M_Guid" -Value $M.Guid
            $item | Add-Member -MemberType NoteProperty -Name "M_OU_DN" -Value "$((($M.DistinguishedName) -split "," -replace "OU=")[(($M.DistinguishedName -split ",").count - 3)..2] -join "/")"
            $item | Add-Member -MemberType NoteProperty -Name "M_OU" -Value $M.Organizationalunit
            $item | Add-Member -MemberType NoteProperty -Name "M_TargetAddress" -Value "$(($M.emailaddresses | where { $_ -match ".mail.onmicrosoft.com" }) -replace "smtp:")"
            $item | Add-Member -MemberType NoteProperty -Name "Identity" -Value $P.Identity
            $item | Add-Member -MemberType NoteProperty -Name "USER_TYPE" -Value $U.Recipienttypedetails
            $item | Add-Member -MemberType NoteProperty -Name "USER" -Value "$($P.user)"
            $item | Add-Member -MemberType NoteProperty -Name "USER_OU" -Value $U.Organizationalunit
            $item | Add-Member -MemberType NoteProperty -Name "USER_TargetAddress" -Value "$(($U.emailaddresses | where { $_ -match ".mail.onmicrosoft.com" }) -replace "smtp:")"
            $item | Add-Member -MemberType NoteProperty -Name "AccessRights" -Value "$({$P.AccessRights})"
            $item | Add-Member -MemberType NoteProperty -Name "SharingPermissionFlags" -Value "$({$P.SharingPermissionFlags})"
            $item | Add-Member -MemberType NoteProperty -Name "USER_SMTP" -Value $U.PrimarySmtpAddress
            $item | Add-Member -MemberType NoteProperty -Name "USER_UPN" -Value $U.userprincipalname
            $item | Add-Member -MemberType NoteProperty -Name "USER_SamAccountName" -Value $U.SamAccountName
            $item | Add-Member -MemberType NoteProperty -Name "USER_OU_DN" -Value "$((($U.DistinguishedName) -split "," -replace "OU=")[(($U.DistinguishedName -split ",").count - 3)..2] -join "/")"

   $Data += $item

         }
      }
   }
   $S =" [MBX] $Num/$count - [Time]"
   $A = "Reading MailboxPermissions [Mailbox Count] ($($Num)/$count) [Mailbox] $($M.PrimarySMTPAddress)"
   Write-Progress -Activity $A -Status $S -PercentComplete (($Num/$count)*100) -SecondsRemaining $($count-$Num) ;
   $Num ++

}

$DATA | Export-Csv c:\Temp\Mailbox_Permissions.csv -NTI -Encoding UTF8 -Force -Delimiter ";"

$DATA.Count

$Orphans | Export-Csv c:\Temp\Mailbox_Permissions_Orphans.csv -NTI -Encoding UTF8 -Force -Delimiter ";"

$Orphans.Count

