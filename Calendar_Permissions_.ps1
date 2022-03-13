$data = @"
email1
email2
"@

$array = $data.Split("`n")
$array = $mbxs.userprincipalname
$array.count

                  $Data2 =@()

foreach ($user in $array) { 

$M = get-EXOMailbox "$($user.Trim())" -ErrorAction SilentlyContinue

$types = "Calendar"

If ($M) {

$D = $M.distinguishedname ; $A = $M.alias ; $U = $M.PrimarySmtpAddress.ToString()

$stats = Invoke-Command {get-EXOMailboxFolderStatistics $A -Folderscope All } -ArgumentList $A

foreach($F in $stats){ 
                  $FN = $A + ':' + $F.FolderPath.Replace('/','\'); 

          write-host "$FN" -F Cyan ; 

$OUT = Invoke-Command {get-EXOMailboxFolderPermission $FN } -ArgumentList $FN

$result = $OUT | where { $_.user.UserType -notlike "anonymous" -and $_.user.UserType -notlike "default" -and $_.user -notlike "NT Authority\SELF**" -and  $_.user -notlike "*S-1-5-21*" }

IF ($result) {

foreach ($P in $result ) {

     $usr = (get-mailbox $P.user.displayname -EA silentlycontinue)
if(!($usr)){  $DG  = (Get-DistributionGroup $P.user.displayname -EA silentlycontinue) } ELSE { $DG = ''    }
if(!($P.user.RecipientPrincipal)){$RcPr = $P.user.RecipientPrincipal } ELSE { $RcPr = ''    }

            $item = New-Object -TypeName PSCustomObject
            $item | Add-Member -MemberType NoteProperty -Name "Identity" -Value $P.Identity
            $item | Add-Member -MemberType NoteProperty -Name "FolderName" -Value $P.FolderName
            $item | Add-Member -MemberType NoteProperty -Name "EMAIL" -Value $user                       
if ($usr) { $item | Add-Member -MemberType NoteProperty -Name "USER" -Value $usr.userprincipalname     }
if ($DG)  { $item | Add-Member -MemberType NoteProperty -Name "Secgroup" -Value $DG.PrimarySmtpAddress }
     ELSE { $item | Add-Member -MemberType NoteProperty -Name "Secgroup" -Value '' }
IF($RcPr) { $item | Add-Member -MemberType NoteProperty -Name "RecipientPrincipal" -Value $RcPr }
            $item | Add-Member -MemberType NoteProperty -Name "user_DisplayName" -Value $P.user.DisplayName
            $item | Add-Member -MemberType NoteProperty -Name "ACCESS" -Value "$(($P.accessrights) -split ', ' -join '|')"
            $item | Add-Member -MemberType NoteProperty -Name "SharingPermissionFlags" -Value ($result.SharingPermissionFlags | Out-String)
            $item | Add-Member -MemberType NoteProperty -Name "UserType" -Value $P.user.UserType
            $item | Add-Member -MemberType NoteProperty -Name "mailbox_DN" -Value $M.distinguishedname
            $item | Add-Member -MemberType NoteProperty -Name "mailbox_UPN" -Value  $M.userprincipalname
            $item | Add-Member -MemberType NoteProperty -Name "mailbox_SMTP" -Value  $M.primarysmtpaddress
            $item | Add-Member -MemberType NoteProperty -Name "mailbox_Alias" -Value $M.alias

               $Data2 += $item

            }
         }
      }
   }
}

$DATA2 | Export-Csv c:\Temp\Calendar_Permissions_source121.csv -NTI -Encoding UTF8 -Force -Delimiter ";"

$DATA2.Count


$userdata = $DATA |where { $_.email -ne $_.user }

$userdata | Export-Csv c:\Temp\ALL_Folder_Permissions_source.csv -NTI -Encoding UTF8 -Force -Delimiter ";"