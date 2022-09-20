$path = "C:\Scripte\Disable_AD_Users\Data"

$Last_CSV = (Get-ChildItem "$Path\Export-LastLoginDates-Cloud_*.csv")[-1].VersionInfo.filename
$lastlogindates = Import-Csv $Last_CSV -Delimiter ";"

#$leer = $lastlogindates | where {$_.lastNonInteractiveSignInDateTime -eq ""}
#$leer.count

$OUs = (Get-ADOrganizationalUnit -Filter * |where {$_ -like "OU=User*OU=DATAGROUP SE,DC=datagroup,DC=local"}).DistinguishedName

#Deaktivierte Benutzer OU
$DecOU = "OU=Inaktive Mitarbeiter,OU=DATAGROUP SE,DC=datagroup,DC=local"

#Anzahl inaktiver Tage
$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))

#Anzahl Tage whencreated
$DaysCreated = 40
$time1 = (Get-Date).Adddays(-($DaysCreated))

$now = Get-Date

$DisableAccs = @()

$OldAccs_abfrage = foreach ($OU in $OUs) {  (Get-ADUser -Server $Server -SearchBase $OU -SearchScope OneLevel -Filter { (LastLogonDate -lt $time) -and (UserAccountControl -ne "514")}  -Properties UserAccountControl,LastLogonDate, whencreated,lastlogontimestamp, logonCount, userprincipalname) }
$OldAccs = $OldAccs_abfrage | select Samaccountname,userprincipalname,distinguishedname,UserAccountControl, LastLogonDate,@{name ="lastlogontimestamp";expression={[datetime]::FromFileTime($_.lastlogontimestamp)}}, whencreated, logonCount

$DisableAccs += $OldAccs

# check 90 days
# check cloudlogin y/n
# check cloudlogin 90 days

$NoLoginAccs_Abfrage =  foreach ($OU in $OUs) { (Get-ADUser -SearchBase $OU -SearchScope OneLevel -Filter {(whencreated -lt $time1) -and (-not ( lastlogontimestamp -like "*"))}  -Properties UserAccountControl, LastLogonDate, whencreated, lastlogontimestamp, logonCount, userprincipalname) }
#$NoLoginAccs

# check 40 Tage whencreated

$NoLoginAccs = $NoLoginAccs_Abfrage | select  Samaccountname,userprincipalname,distinguishedname,UserAccountControl, LastLogonDate,@{name ="lastlogontimestamp";expression={[datetime]::FromFileTime($_.lastlogontimestamp)}}, whencreated, logonCount

$DisableAccs += $NoLoginAccs

$DisableAccs.count

 #  (Get-aduser $OldAccs -Properties userprincipalname).userprincipalname

#$UPNS = $OldAccs | % { Get-aduser $_ -Properties userprincipalname,mail }

$AllData = @()

foreach ($Acc in $DisableAccs) {

$item = $null
$deactivate = $null
$onprem_Logonstatus = $null
$Cloud_Logonstatus = $null

# $Acc.lastlogontimestamp

If ([string]::IsNullOrEmpty($Acc.lastlogontimestamp)) {

($Acc.whencreated -gt $time1)) { $onprem_status = "NEU_inaktiv" }

If ([string]::IsNullOrEmpty($Acc.lastlogontimestamp) -and ($Acc.whencreated -lt $time1)) { $onprem_status = "NEU_expired" } ## OnPREM kein Login und älter als 40 Tage

If (!([string]::IsNullOrEmpty($Acc.lastlogontimestamp)) -and ($Acc.LastLogonDate -lt $time)) { $onprem_status = "Inaktiv_expired" }  ## OnPREM Älter als 90 Tage

If (!([string]::IsNullOrEmpty($Acc.lastlogontimestamp)) -and ($Acc.LastLogonDate -gt $time)) { $onprem_status = "Aktiv" }  ## OnPREM Älter als 90 Tage

$item = $lastlogindates.where({$_.userPrincipalName -eq $Acc.userprincipalname })

If ($item -and ([string]::IsNullOrEmpty($item.lastSignInDateTime))) { $cloud_status = "Synced_NoCloudLogin" }

If ($item -and ($item.lastSignInDateTime -ge $time))      { $cloud_status = "Synced_and_Login" }

If ($item -and ($item.lastSignInDateTime -lt $time))      { $cloud_status = "Synced_and_Expired" }

If (!($item))      { $cloud_status = "No_Clouduser_found" }

if ($item[1])      { $cloud_status = "Duplikat" ; $item = $item[0] }

###
If ($item.lastSignInDateTime -and (!($Acc.LastLogonDate))) { $cloud_Data = "Synced_CloudOnly" }###


if ($onprem_status -match "aktiv" -or $onprem_status -eq "Inaktiv_expired" -or $onprem_status -eq "Aktiv" ) {

# 01.01.1601 01:00:00
# Onprem

If ([string]::IsNullOrEmpty($Acc.lastlogontimestamp)) {

  [DateTime]$onprem_LastLogonDate = get-date $Acc.LastLogonDate -ErrorAction Stop
                [String]$onprem_LogOnAge = ($now - $onprem_LastLogonDate).Days
              $onprem_lastlogontimestamp = $Acc.lastlogontimestamp

                          $Error_String1 = ""

                           } catch { $Error[0].Exception

           [String]$onprem_LastLogonDate =  [String]($Acc.LastLogonDate)
                [String]$onprem_LogOnAge =  "invalid"   
    [DateTime]$onprem_lastlogontimestamp =  [String]($Acc.lastlogontimestamp)

                          $Error_String1 = "$($Error[0].Exception.Message)"

                           }

} Elseif ($onprem_status -match "NEU") {

# Onprem - No Data       
                       $onprem_LastLogonDate =  get-date "01.01.0001 01:00:00"
                    [String]$onprem_LogOnAge =  "Never"


} Else {

# Onprem - No Data       
                       $onprem_LastLogonDate =  get-date "01.01.0001 01:00:00"
                    [String]$onprem_LogOnAge =  "No_confirmed_status"


}

if ($cloud_status -eq "Synced_and_Login" -or $cloud_status -eq "Synced_CloudOnly" -or $cloud_status -eq "Synced_and_Expired" -or $cloud_status -eq "Duplikat" ) {

# Cloud  

IF ( [string]::IsNullOrEmpty($item.lastSignInDateTime)) {
           [String]$Cloud_lastSignInDate =  [String]($item.lastSignInDateTime)
            [String]$Cloud_lastSignInAge = "Never"
                          $Error_String2 = ""
} ELSE {

 Try {   [DateTime]$Cloud_lastSignInDate = get-date $item.lastSignInDateTime -ErrorAction Stop
            [String]$Cloud_lastSignInAge = ($now - $Cloud_lastSignInDate).Days
                          $Error_String2 = ""

                           } catch { $Error[0].Exception

           [String]$Cloud_lastSignInDate =  [String]($item.lastSignInDateTime)
            [String]$Cloud_lastSignInAge = "invalid"  
                          $Error_String2 = "$($Error[0].Exception.Message)"
                           }
       }
 # Cloud N/I

IF ( [string]::IsNullOrEmpty($item.lastNonInteractiveSignInDateTime)) {

        [String]$Cloud_lastSignInDate_NI =  [String]($item.lastNonInteractiveSignInDateTime)
         [String]$Cloud_lastSignInAge_NI = "Never"
                          $Error_String3 = ""
} ELSE {

Try { [DateTime]$Cloud_lastSignInDate_NI = get-date $item.lastNonInteractiveSignInDateTime -ErrorAction Stop
            [int]$Cloud_lastSignInAge_NI = ($now - $Cloud_lastSignInDate).Days
                          $Error_String3 = ""

                          } catch { $Error[0].Exception

        [String]$Cloud_lastSignInDate_NI =  [String]($item.lastNonInteractiveSignInDateTime)
         [String]$Cloud_lastSignInAge_NI = "invalid"
                          $Error_String3 = "$($Error[0].Exception.Message)"
                          }
      }
}

Elseif ($cloud_status -eq "No_Clouduser_found" -or $cloud_status -eq "Synced_NoCloudLogin") {

# Cloud - No Data  

                   $Cloud_lastSignInDate = get-date "01.01.0001 01:00:00" # [String]($item.lastSignInDateTime)
            [String]$Cloud_lastSignInAge = "NONE"

# Cloud N/I- No Data 

                $Cloud_lastSignInDate_NI = get-date "01.01.0001 01:00:00" # [String]($item.lastNonInteractiveSignInDateTime)
         [String]$Cloud_lastSignInAge_NI = "NONE"

}

Else {

# Cloud - No Data  

                   $Cloud_lastSignInDate = [String]($item.lastSignInDateTime)
            [String]$Cloud_lastSignInAge = "No_confirmed_status"

# Cloud N/I- No Data 

                $Cloud_lastSignInDate_NI = [String]($item.lastNonInteractiveSignInDateTime)
         [String]$Cloud_lastSignInAge_NI = "No_confirmed_status"

}


$OU =(($Acc.DistinguishedName) -split "," -replace "OU=")[(($Acc.DistinguishedName -split ",").count - 3)..1] -join "/" 

## Deaktivierungsstatus
#If (($onprem_LastLogonDate -lt $time -or $onprem_LastLogonDate -eq "NONE")  -and ($Cloud_lastSignInDate_NI -lt $time -or $Cloud_lastSignInDate_NI -eq "NONE"))
If (($cloud_status -eq "No_Clouduser_found" -or $cloud_status -eq "Synced_NoCloudLogin" -or $cloud_status -eq "Synced_and_Expired") -and ($onprem_status -eq "Inaktiv_expired" -or $onprem_status -eq "NEU_expired"))
{

$deactivate = $true

If ($run_inbackground_noninteractive -eq $false ) {
#Write-host "$OU | USER wird deaktiviert:  $($item.displayName) [$($item.userPrincipalName)]" -NoNewline -ForegroundColor Yellow
#Write-host "  Alter: Onprem-Cloud-Cloud(N/I)[$onprem_LogOnAge | $Cloud_lastSignInAge | $Cloud_lastSignInAge_NI ]"  -ForegroundColor Cyan
}
} else {

$deactivate = $false
If ($run_inbackground_noninteractive -eq $false ) {
#Write-host "$OU | USER ist noch aktiv:  $($item.displayName) [$($item.userPrincipalName)]"  -NoNewline -ForegroundColor green
#Write-host "  Alter: Onprem-Cloud-Cloud(N/I)[$onprem_LogOnAge | $Cloud_lastSignInAge | $Cloud_lastSignInAge_NI ]"  -ForegroundColor Cyan
}
}



# lastlogontimestamp

        $Data = New-Object -TypeName PSObject      
        $Data | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $Acc.SamAccountName
        $Data | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $Acc.userprincipalname
        $Data | Add-Member -MemberType NoteProperty -Name Name -Value $Acc.Name
        $Data | Add-Member -MemberType NoteProperty -Name UserAccountControl -Value $Acc.UserAccountControl
        $Data | Add-Member -MemberType NoteProperty -Name Distinguishedname -Value $Acc.Distinguishedname
        $Data | Add-Member -MemberType NoteProperty -Name OU -Value $OU
     
        $Data | Add-Member -MemberType NoteProperty -Name deactivate -Value $deactivate
        $Data | Add-Member -MemberType NoteProperty -Name Data_Status -Value $cloud_Data

        $Data | Add-Member -MemberType NoteProperty -Name onprem_LastLogonDate -Value $onprem_LastLogonDate
        $Data | Add-Member -MemberType NoteProperty -Name onprem_LogOnAge -Value $onprem_LogOnAge

If ($Acc.lastlogontimestamp) {
        $Data | Add-Member -MemberType NoteProperty -Name onprem_lastlogontimestamp -Value $Acc.lastlogontimestamp
} Else{ $Data | Add-Member -MemberType NoteProperty -Name onprem_lastlogontimestamp -Value "Never" }

        $Data | Add-Member -MemberType NoteProperty -Name logonCount -Value $Acc.logonCount

        $Data | Add-Member -MemberType NoteProperty -Name Cloud_lastSignInDate -Value $Cloud_lastSignInDate
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_lastSignInAge -Value $Cloud_lastSignInAge

        $Data | Add-Member -MemberType NoteProperty -Name Cloud_lastSignInDate_NI -Value $Cloud_lastSignInDate_NI
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_lastSignInAge_NI -Value $Cloud_lastSignInAge_NI

        $Data | Add-Member -MemberType NoteProperty -Name Error_String_1 -Value $Error_String1
        $Data | Add-Member -MemberType NoteProperty -Name Error_String_2 -Value $Error_String2
        $Data | Add-Member -MemberType NoteProperty -Name Error_String_3 -Value $Error_String3

$AllData += $Data

}

$AllData.count

$Path = "C:\Scripte\Disable_AD_Users\Data"

$TS = (get-date -Format yyyy-MM-dd_HH.mm).ToString()

$AllData | Export-CSV -Path "$Path\Disable_Entwurf_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$OldAccs | Export-CSV -Path "$Path\OldAccs_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$NoLoginAccs | Export-CSV -Path "$Path\NoLoginAccs_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$DisableAccs | Export-CSV -Path "$Path\DisableAccs_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation



Compress-Archive -Path "$Path\Disable_Entwurf_$TS.CSV" -DestinationPath "$Path\Disable_Entwurf_$TS.zip" -Force # Zip Logs
#Invoke-Item $Path # open file manager


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

######### Fill Out Variables #############################

$FROM = "Test01@datagroup.de"

#$TO = "marius.mueller@datagroup.de", "jan.huebener@datagroup.de" #, "third@recipient.com"

$TO ="jan.huebener@datagroup.de" #, "third@recipient.com"

$attachment = "$Path\Disable_Entwurf_$TS.zip"

##########################################################

#$password = ConvertTo-SecureString "******" -AsPlainText -Force
#$credential = New-Object System.Management.Automation.PSCredential ("$FROM", $password)
#Get-Credential "Test01@datagroup.de" | Export-Clixml "C:\Scripte\Disable_AD_Users\cred\Credentials.xml"

$Credential = Import-Clixml -Path "C:\Scripte\Disable_AD_Users\cred\Credentials.xml"

$smtpserver = "smtp.datagroup.local"
$Port = "25"

#$smtpserver = "smtp.office365.com"
#$Port = "587"
##########################################################
$Subject = "Bericht deaktivierter Benutzer (automatiert)"

####################  choose 1/2/3  ######################


$head = "<style>
td {width:100px; max-width:300px; background-color:lightgrey;}
table {width:100%;}
th {font-size:14pt;background-color:yellow;}
</style>

<title>Report zum Deaktivierungsscript</title>"

$datum = Get-Date -Format "dd.MM.yyyy, HH:mm"

$body = $AllData |where { $_.deactivate -eq $true } | ConvertTo-Html -Head $head -PreContent "<h1>Report erzeugt von $env:USERNAME am $datum</h1>"

##########################################################
$message = new-object System.Net.Mail.MailMessage
$message.From = $FROM
$message.To.Add($TO)
$message.IsBodyHtml = $True
$message.Subject = $Subject
$attach = new-object Net.Mail.Attachment($attachment)
$message.Attachments.Add($attach)
$message.body = $body

$smtp = new-object Net.Mail.SmtpClient($SmtpServer, 25)
$smtp.EnableSsl = $false
$smtp.Credentials = $credential;
$smtp.Send($message)
########################################################## 
