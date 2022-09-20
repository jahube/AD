## Erstellt von Marius Müller ##

#Import AD Modul
import-module activedirectory 

#Logpfad
#$Log = "C:\Scripte\Disable_AD_Users\Disable_Inactive_AD_Users.txt"
$Log = "C:\Scripte\Disable_AD_Users\Disable_Inactive_AD_Users.txt"

#Festlegen der Zeit im Log
$logtime = Get-Date -Format "[dd.MM.yy HH:mm]:"

#Server festlegen
#Definiert die möglichen DCs
$DCs = "DGDC-DC-01", "DGDC-DC-02", "DGDC-DC-03"

#Leert den Server Parameter
$Server = $null

$AllData = @()

$Alt_Inaktiv__Deaktiviert = @()

$Neu_Inaktiv_Deaktiviert = @()

$CloudOnly_Aktiv = @()

#Checkt alle vorhandenen DCs auf Online Status
$DCs | foreach {
            if (Test-Connection $_ -Count 1 -Quiet) 
            {
                #Wenn der Server verfügbar ist, wird dieser verwendet
                if(-not($GetObject))
                {
                    #Wenn Server Variable leer, dann setzten
                    if(!$Server)
                    {
                        $Server = $($_)
                        Write-Host -ForegroundColor green "$Server wird aktuell verwendet."
                        Write-Output "$logtime [Information] $Server wird für die Ausführung des Scripts verwendet." >> $Log
                    }
                    
                }
                
            }

            #Wenn der Server nicht verfügbar ist, wird in Log geschrieben
            else
            {
                    if(-not($GetObject))
                    {
                    write-host -ForegroundColor Red "$_ ist nicht verfügbar!"
                    Write-Output "$logtime [Information] Der Server $_ ist nicht verfügbar!" >> $Log
                    }
            }
        }


#Alle Benutzer OUs aus dem AD auslesen und in das Array $OUs schreiben

$OUs = @()

$getous = (Get-ADOrganizationalUnit -Filter * | Where-Object -FilterScript {$PSItem.distinguishedname -like "*OU=DATAGROUP SE,DC=datagroup,DC=local"}).DistinguishedName
foreach($getou in $getous)
{
    if($getou -like "OU=User*")
    {
        $OUs += $getou

    }
}

#Deaktivierte Benutzer OU
$DecOU = "OU=Inaktive Mitarbeiter,OU=DATAGROUP SE,DC=datagroup,DC=local"

#Anzahl inaktiver Tage
$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))

#Anzahl Tage whencreated
$DaysCreated = 40
$time1 = (Get-Date).Adddays(-($DaysCreated))

$Now = (Get-Date)

Write-Output "$logtime [Info] Alle Accounts die mehr als $DaysInactive Tage nicht angemeldet waren, werden deaktiviert!" >> $Log
Write-Output "$logtime [Info] Alle Accounts die vor mehr als $DaysCreated Tage angelegt und nicht angemeldet waren, werden deaktiviert!" >> $Log

#Prüfung Cloud LogOnDate Datei
$Last_CSV = (Get-ChildItem "C:\Scripte\Disable_AD_Users\Data\Export-LastLoginDates-Cloud_*.csv")[-1]
$Last_CSV.VersionInfo.filename
[DATETIME]$CloudCSV_CreationDate = $Last_CSV.CreationTime
$CloudCSV = Import-Csv $Last_CSV -Delimiter ";"

#####################

If ($CloudCSV_Creationdate -gt ((Get-Date).AddDays(-7)))
{
    $CloudCSV_Valid = $true
    Write-Output "$logtime [Info] Cloud File up to date!" >> $Log
}
Else
{
    $CloudCSV_Valid = $false
    #SENDMAIL Cloud File not up to date.
    Write-Output "$logtime [Info] Cloud File not up to date!" >> $Log
}
#####################


#Prüfung für jede OU, und ob Account enabled ist und Tage > Daysinactive
foreach($OU in $OUs)
    {

        Write-Output "$logtime [Info] $OU wird verarbeitet." >> $Log

        #Alle Accounts der OU werden geprüft, ob bereits deaktivert und das LogOnDate größer $DaysInactive ist (LastLogonDate Never wird ignoriert!)
        $OldAccs = Get-ADUser -Server $Server -SearchBase $OU -SearchScope OneLevel -Filter {(LastLogonDate -lt $time) -and (UserAccountControl -ne "514")}  -Properties UserAccountControl, LastLogonDate, whencreated, lastlogontimestamp, logonCount, userprincipalname

        
            #Deaktivierung und verschiebung der Accounts in vorgesehene OU
            foreach ($NAcc_item in $OldAccs)

            {   #Ausführung 
                $logtime = Get-Date -Format "[dd.MM.yy HH:mm]:"
 ##### AD User Properties
                $NAcc = $NAcc_item.SamAccountname
                $Acc = $NAcc_item.SamAccountname
                $NAcc_UPN = $NAcc_item.userprincipalname
                $NAcc_whencreated = $NAcc_item.whencreated

                $NAcc_lastlogon = $NAcc_item.LastLogonDate
             IF ($NAcc_lastlogon -eq $null -or $NAcc_lastlogon.length -eq 0) { $NAcc_lastlogon = [DateTime]("01.01.1601 01:00:00") }

                $AD_Logon_Age = try { (New-TimeSpan -start $NAcc_lastlogon -end $now -ErrorAction Stop).Days} 
                              catch { return [int]("9999") }
 ##### Cloud User Properties
                $Cloud_Acc = $CloudCSV.where({$_.userPrincipalName -eq $NAcc_UPN })

                $Nacc_OU =(($NAcc_item.DistinguishedName) -split "," -replace "OU=")[(($NAcc_item.DistinguishedName -split ",").count - 3)..1] -join "/" 


 ##### Cloud Check
                #If ($Cloud_Acc -and (!( [string]::IsNullOrEmpty($Cloud_Acc.lastSignInDateTime))))

                If ($Cloud_Acc -and ($Cloud_Acc.lastSignInDateTime -ne $null) -and ($Cloud_Acc.lastSignInDateTime.length -ne 0))
                {

                    Write-Output "$logtime [Info] $Acc has Cloud Logindate!" >> $Log
                    Write-Host "$logtime [Info] $Acc has Cloud Logindate!"
                    $CloudLogin = $true
                }
                Else
                {

                    Write-Output "$logtime [Info] $Acc has NO Cloud Logindate!" >> $Log
                    Write-Host "$logtime [Info] $Acc has NO Cloud Logindate!"
                    $CloudLogin = $false
                }
##

                If ($CloudLogin)
                {

                $dateString = $Cloud_Acc.lastSignInDateTime
                $format = "yyyy-MM-ddTHH:mm:ssZ"
                [ref]$parsedDate = get-date
                $parsed = [DateTime]::TryParseExact($dateString, $format,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,$parseddate)

                     if($parsed)
                     {
                       #  write "$dateString is valid"

                    # $parseddate.Value.ToString("yyyy-MM-ddTHH:mm:ssZ")

                     $cloudage = try { (New-TimeSpan -start $parseddate.Value -end $now -ErrorAction Stop).Days} 
                               catch { return [int]("999") }
                     }
                     Else 
                     { 
                     $cloudage = [int]("9999") 
                     }

                }
                Else 
                { 
                $cloudage = [int]("99999") 
                }



#####
        $Data = New-Object -TypeName PSObject      
        $Data | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $NAcc_item.SamAccountName
        $Data | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $NAcc_item.userprincipalname
        $Data | Add-Member -MemberType NoteProperty -Name Name -Value $NAcc_item.Name
        $Data | Add-Member -MemberType NoteProperty -Name Reason -Value "Expired_90_Tage"
        $Data | Add-Member -MemberType NoteProperty -Name UserAccountControl -Value $NAcc_item.UserAccountControl
        $Data | Add-Member -MemberType NoteProperty -Name Distinguishedname -Value $NAcc_item.Distinguishedname
        $Data | Add-Member -MemberType NoteProperty -Name AD_logonCount -Value $NAcc_item.logonCount
        $Data | Add-Member -MemberType NoteProperty -Name OU -Value $Nacc_OU

If ($NAcc_item.LastLogonDate -ne $null -and $NAcc_item.LastLogonDate.length -ne 0){
        
        [String]$NAcc_item_LastLogonDate = [String]($NAcc_item.LastLogonDate)
        $Data | Add-Member -MemberType NoteProperty -Name AD_LastLogonDate -Value $NAcc_item_LastLogonDate

         $AD_Logon_Age =   try { (New-TimeSpan -start $NAcc_lastlogon -end $now -ErrorAction Stop).Days} 
                         catch { return [int]("9999") }
        $AD_Logon_Age = [String]$AD_Logon_Age

        $Data | Add-Member -MemberType NoteProperty -Name AD_Logon_Age -Value $AD_Logon_Age

} Else{ $Data | Add-Member -MemberType NoteProperty -Name AD_LastLogonDate -Value "Never"
        $Data | Add-Member -MemberType NoteProperty -Name AD_Logon_Age -Value "Never"

 }

        $Data | Add-Member -MemberType NoteProperty -Name CloudLogin -Value $CloudLogin
 IF ($CloudLogin -and $parsed) {

[String]$Cloud_LastLogonDate = $parseddate.Value
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LastLogonDate -Value $Cloud_LastLogonDate

        [String]$Cloud_LogOnAge = [String]($cloudage)
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LogOnAge -Value $Cloud_LogOnAge
} Else{ 
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LastLogonDate -Value "Never"
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LogOnAge -Value "Never"
}


                ##Prüfen ob der Account ein Service Benutzer ist, wenn ja dann Error
                If ($Acc -like "SVC_*") 
                {
                    Write-Output "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!" >> $Log
                    Write-Host "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!"
                    $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "SRV"

 
                }

                ##Ausnahme für DGSOPS.Flow (Service Account für Cloud Flows)
                elseif($Acc -eq "DGSOPS.Flow")
                {
                    Write-Output "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!" >> $Log
                    Write-Host "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!"
                    $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Flow"

                }
                
                elseif($cloudage -lt 90)

                {
                    Write-Output "$logtime [Cloud - LastLogOn] $Acc active in Cloud $cloudage days ago!" >> $Log
                    Write-Host "$logtime [Cloud - LastLogOn] $Acc active in Cloud $cloudage days ago!"
                    $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Cloud_aktiv"

                    $CloudOnly_Aktiv += $Data
                }
                ##Wenn der Account kein Service Benutzer ist, wird er deaktiviert
                Else
                {
                    #Log Information für Benutzer Deaktivierung
                    #Write-Output "$logtime [Try - LastLogOn] $Acc wird deaktiviert..." >> $Log
                    Write-Host "$logtime [Try - LastLogOn] $Acc wird deaktiviert..."
                    
                                        #Deaktivierung des Accounts
                    Disable-ADAccount -Server $Server -identity $Acc

                    <#
                    #Entfernen der Lizenzgruppen
                    $groups = (Get-ADPrincipalGroupMembership -Identity $Acc).Name
                    $licgroups = $groups | Where { $_ -like "DGSE_LIC_M365_*"}
                    foreach($licgroup in $licgroups)
                    {
                        Write-Output "$logtime [Success - GroupMatch] Removing $Acc from Group: $licgroup" >> $Log
                        Remove-ADGroupMember -Identity $licgroup -Members $Acc -Confirm:$false
                    }
                    #>
                    
                    #Verschiebung des Accounts
                    $AccDN = (Get-ADUser -Identity $Acc).DistinguishedName
                    Move-ADObject -Server $Server -Identity $AccDN -TargetPath $DecOU
                    
                    #Überprüfung auf deaktivierung und verschiebung des Accounts 
                    $AccCheck = Get-AdUser -Server $Server -identity $Acc -Properties Enabled, DistinguishedName, LastLogonDate,whencreated
                    $enabledcheck = $AccCheck.Enabled
                    $dncheck = $AccCheck.DistinguishedName
                    $ulastlogon = $AccCheck.LastLogonDate
                    
                    IF ($ulastlogon -eq $null -or $ulastlogon.length -eq 0) { $ulastlogon = [DateTime]("01.01.1601 01:00:00") }

                    $ucreated = $AccCheck.whencreated

                    #Bei erfolgreicher Verschiebung wird folgendes Ausgegeben:
                    if ($enabledcheck -eq $false -and $dncheck -like "*$DecOU")
                    {
                        Write-Output "$logtime [Success - LastLogOn] $Acc wurde deaktiviert und verschoben. LastLogOn: $ulastlogon" >> $Log
                        Write-Host "$logtime [Success - LastLogOn] $Acc wurde deaktiviert und verschoben. LastLogOn: $ulastlogon"
                        $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Erfolgreich"
                    }
                    
                    #Wenn Operation auf Fehler stösst:
                    else
                    {
                        Write-Output "$logtime [ERROR - LastLogOn] $Acc wurde nicht deaktiviert und verschoben. Bitte prüfen! LastLogOn: $ulastlogon" >> $Log
                        Write-Host "$logtime [ERROR - LastLogOn] $Acc wurde nicht deaktiviert und verschoben. Bitte prüfen! LastLogOn: $ulastlogon"
                        $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Fehlerhaft"

                    }

                       $Alt_Inaktiv__Deaktiviert += $Data
                }

                $AllData += $Data
            }



        #Alle Accounts der OU werden geprüft, ob LastLogonDate Never entspricht und der Account vor mehr als 40 Tagen erstellt wurde
        $NoLoginAccs = Get-ADUser -SearchBase $OU -SearchScope OneLevel -Filter {(whencreated -lt $time1) -and (-not ( lastlogontimestamp -like "*"))}  -Properties UserAccountControl, LastLogonDate, whencreated, lastlogontimestamp, logonCount, userprincipalname

            #Info mit Benutzern ohne LastLogonDate ins Logfile
            foreach ($NAcc_item in $NoLoginAccs)
            {   #Ausführung 
                $logtime = Get-Date -Format "[dd.MM.yy HH:mm]:"
 ##### AD User Properties
                $NAcc = $NAcc_item.SamAccountname
                $Acc = $NAcc_item.SamAccountname
                $NAcc_UPN = $NAcc_item.userprincipalname
                $NAcc_whencreated = $NAcc_item.whencreated

                $NAcc_lastlogon = $NAcc_item.LastLogonDate

             IF ($NAcc_lastlogon -eq $null -or $NAcc_lastlogon.length -eq 0) { $NAcc_lastlogon = [DateTime]("01.01.1601 01:00:00") }

 ##### Cloud User Properties

                $Cloud_Acc = $CloudCSV.where({$_.userPrincipalName -eq $NAcc_UPN })

                $Nacc_OU =(($NAcc_item.DistinguishedName) -split "," -replace "OU=")[(($NAcc_item.DistinguishedName -split ",").count - 3)..1] -join "/" 


 ##### Cloud Check
                #If ($Cloud_Acc -and (!( [string]::IsNullOrEmpty($Cloud_Acc.lastSignInDateTime))))

                If ($Cloud_Acc -and ($Cloud_Acc.lastSignInDateTime -ne $null) -and ($Cloud_Acc.lastSignInDateTime.length -ne 0))
                {

                    Write-Output "$logtime [Info] $Acc has Cloud Logindate!" >> $Log
                    Write-Host "$logtime [Info] $Acc has Cloud Logindate!"
                    $CloudLogin = $true
                }
                Else
                {

                    Write-Output "$logtime [Info] $Acc has NO Cloud Logindate!" >> $Log
                    Write-Host "$logtime [Info] $Acc has NO Cloud Logindate!"
                    $CloudLogin = $false
                }
##

                If ($CloudLogin)
                {

                $dateString = $Cloud_Acc.lastSignInDateTime
                $format = "yyyy-MM-ddTHH:mm:ssZ"
                [ref]$parsedDate = get-date
                $parsed = [DateTime]::TryParseExact($dateString, $format,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,$parseddate)

                     if($parsed)
                     {
                       #  write "$dateString is valid"

                    # $parseddate.Value.ToString("yyyy-MM-ddTHH:mm:ssZ")

                     $cloudage = try { (New-TimeSpan -start $parseddate.Value -end $now -ErrorAction Stop).Days} 
                               catch { return [int]("999") }
                     }
                     Else 
                     { 
                     $cloudage = [int]("9999") 
                     }

                }
                Else 
                { 
                $cloudage = [int]("99999") 
                }

#####
        $Data = New-Object -TypeName PSObject      
        $Data | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $NAcc_item.SamAccountName
        $Data | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $NAcc_item.userprincipalname
        $Data | Add-Member -MemberType NoteProperty -Name Name -Value $NAcc_item.Name
        $Data | Add-Member -MemberType NoteProperty -Name Reason -Value "Neu_40_Tage"
        $Data | Add-Member -MemberType NoteProperty -Name UserAccountControl -Value $NAcc_item.UserAccountControl
        $Data | Add-Member -MemberType NoteProperty -Name Distinguishedname -Value $NAcc_item.Distinguishedname
        $Data | Add-Member -MemberType NoteProperty -Name AD_logonCount -Value $NAcc_item.logonCount
        $Data | Add-Member -MemberType NoteProperty -Name OU -Value $Nacc_OU

If ($NAcc_item.LastLogonDate -ne $null -and $NAcc_item.LastLogonDate.length -ne 0){
        
        [String]$NAcc_item_LastLogonDate = [String]($NAcc_item.LastLogonDate)
        $Data | Add-Member -MemberType NoteProperty -Name AD_LastLogonDate -Value $NAcc_item_LastLogonDate

         $AD_Logon_Age =   try { (New-TimeSpan -start $NAcc_lastlogon -end $now -ErrorAction Stop).Days} 
                         catch { return [int]("9999") }
        $AD_Logon_Age = [String]$AD_Logon_Age

        $Data | Add-Member -MemberType NoteProperty -Name AD_Logon_Age -Value $AD_Logon_Age

} Else{ $Data | Add-Member -MemberType NoteProperty -Name AD_LastLogonDate -Value "Never"
        $Data | Add-Member -MemberType NoteProperty -Name AD_Logon_Age -Value "Never"

 }

        $Data | Add-Member -MemberType NoteProperty -Name CloudLogin -Value $CloudLogin
 IF ($CloudLogin -and $parsed) {

[String]$Cloud_LastLogonDate = $parseddate.Value
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LastLogonDate -Value $Cloud_LastLogonDate

        [String]$Cloud_LogOnAge = [String]($cloudage)
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LogOnAge -Value $Cloud_LogOnAge
} Else{ 
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LastLogonDate -Value "Never"
        $Data | Add-Member -MemberType NoteProperty -Name Cloud_LogOnAge -Value "Never"
}

            
                ##Prüfen ob der Account ein Service Benutzer ist, wenn ja dann Error
                If ($NAcc -like "SVC_*") 
                {
                    Write-Output "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!" >> $Log
                    Write-Host "$logtime [ERROR / SVC - LastLogOn] $Acc is a Service Account!"
                    $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "SRV"

 
                }

                elseif($cloudage -lt 90)

                {
                    Write-Output "$logtime [Cloud - LastLogOn] $Acc active in Cloud $cloudage days ago!" >> $Log
                    Write-Host "$logtime [Cloud - LastLogOn] $Acc active in Cloud $cloudage days ago!"
                    $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Cloud_aktiv"
                    $CloudOnly_Aktiv += $Data
                }

                ##Wenn der Account kein Service Benutzer ist, wird er deaktiviert
                Else
                {
                    #Log Information für Benutzer Deaktivierung
                    #Write-Output "$logtime [Try - NoLogOn] $NAcc wird deaktiviert..." >> $Log
                    Write-Host "$logtime [Try - NoLogOn] $NAcc wird deaktiviert..."
                    
                                        #Deaktivierung des Accounts
                    Disable-ADAccount -Server $Server -identity $NAcc

                    <#
                    #Entfernen der Lizenzgruppen
                    $groups = (Get-ADPrincipalGroupMembership -Identity $NAcc).Name
                    $licgroups = $groups | Where { $_ -like "DGSE_LIC_M365_*"}
                    foreach($licgroup in $licgroups)
                    {
                        Write-Output "$logtime [Success - GroupMatch] Removing $NAcc from Group: $licgroup" >> $Log
                        Remove-ADGroupMember -Identity $licgroup -Members $NAcc -Confirm:$false
                    }
                    #>
                    
                    #Verschiebung des Accounts
                    $AccDN = (Get-ADUser -Identity $NAcc).DistinguishedName
                    Move-ADObject -Server $Server -Identity $AccDN -TargetPath $DecOU
                    
                    #Überprüfung auf deaktivierung und verschiebung des Accounts 
                    $AccCheck = Get-AdUser -Server $Server -identity $NAcc -Properties Enabled, DistinguishedName, LastLogonDate,whencreated
                    $enabledcheck = $AccCheck.Enabled
                    $dncheck = $AccCheck.DistinguishedName
                    $ulastlogon = $AccCheck.LastLogonDate

                    IF ($ulastlogon -eq $null -or $ulastlogon.length -eq 0) { $ulastlogon = [DateTime]("01.01.1601 01:00:00") }

                    $ucreated = $AccCheck.whencreated
                    
                                  
                    #Bei erfolgreicher Verschiebung wird folgendes Ausgegeben:
                    if ($enabledcheck -eq $false -and $dncheck -like "*$DecOU")
                    {
                        Write-Output "$logtime [Success - NoLogOn] $NAcc wurde deaktiviert und verschoben. Angelegt: $ucreated" >> $Log
                        Write-Host "$logtime [Success - NoLogOn] $NAcc wurde deaktiviert und verschoben. Angelegt: $ucreated"
                        $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Erfolgreich"

                    }
                    
                    #Wenn Operation auf Fehler stösst:
                    else
                    {
                        Write-Output "$logtime [ERROR - NoLogOn] $NAcc wurde nicht deaktiviert und verschoben. Bitte prüfen! Angelegt: $ucreated" >> $Log
                        Write-Host "$logtime [ERROR - NoLogOn] $NAcc wurde nicht deaktiviert und verschoben. Bitte prüfen! Angelegt: $ucreated"
                        $Data | Add-Member -MemberType NoteProperty -Name Deaktivierung -Value "Fehlerhaft"

                    }
                    $Neu_Inaktiv_Deaktiviert += $Data
                }

                $AllData += $Data

            }
    }


$Path = "C:\Scripte\Disable_AD_Users\Data"

$TS = (get-date -Format yyyy-MM-dd_HH.mm).ToString()

$AllData | Export-CSV -Path "$Path\AllData_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$Alt_Inaktiv__Deaktiviert | Export-CSV -Path "$Path\Alt_Inaktiv__Deaktiviert_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$Neu_Inaktiv_Deaktiviert | Export-CSV -Path "$Path\Neu_Inaktiv_Deaktiviert_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

$CloudOnly_Aktiv | Export-CSV -Path "$Path\CloudOnly_Aktiv_$TS.CSV" -Delimiter ";" -Encoding UTF8 -Force -NoTypeInformation

Compress-Archive -Path "$Path\AllData_$TS.CSV" -DestinationPath "$Path\AllData_$TS.zip" -Force # Zip Logs
Compress-Archive -Path "$Path\Alt_Inaktiv__Deaktiviert_$TS.CSV" -DestinationPath "$Path\Alt_Inaktiv__Deaktiviert_$TS.zip" -Force # Zip Logs
Compress-Archive -Path "$Path\Neu_Inaktiv_Deaktiviert_$TS.CSV" -DestinationPath "$Path\Neu_Inaktiv_Deaktiviert_$TS.zip" -Force # Zip Logs
Compress-Archive -Path "$Path\CloudOnly_Aktiv_$TS.CSV" -DestinationPath "$Path\CloudOnly_Aktiv_$TS.zip" -Force # Zip Logs

#Invoke-Item $Path # open file manager


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

######### Fill Out Variables #############################

$FROM = "SRV_Report@datagroup.de"

#$TO = "marius.mueller@datagroup.de", "jan.huebener@datagroup.de", "simon.gienger@datagroup.de", "daniel.berndt@datagroup.de"

#$TO = "jan@jan.red", "jan.huebener@datagroup.de", "jan.hubener@outlook.com"

$TO = "jan.huebener@datagroup.de"

#$TO ="VL.DG.Stuttgart.Operations-KIT@datagroup.de" #, "third@recipient.com"

$attachment = "$Path\Disable_Entwurf_$TS.zip"

$attachment1 = "$Path\AllData_$TS.zip"
$attachment2 = "$Path\Alt_Inaktiv__Deaktiviert_$TS.zip"
$attachment3 = "$Path\Neu_Inaktiv_Deaktiviert_$TS.zip"
$attachment4 = "$Path\CloudOnly_Aktiv_$TS.zip"


##########################################################

#$password = ConvertTo-SecureString "******" -AsPlainText -Force
#$credential = New-Object System.Management.Automation.PSCredential ("$FROM", $password)
#Get-Credential "SRV_Report@datagroup.de" | Export-Clixml "C:\Scripte\Disable_AD_Users\cred\SRV_Report.xml"

$Credential = Import-Clixml -Path "C:\Scripte\Disable_AD_Users\cred\SRV_Report.xml"

$smtpserver = "smtp.datagroup.local"
$Port = "25"

#$smtpserver = "smtp.office365.com"
#$Port = "587"
##########################################################
$Subject = "Bericht deaktivierter Benutzer (automatisiert)"

####################  choose 1/2/3  ######################


$head = "<style>
td {width:100px; max-width:300px; background-color:lightgrey;}
table {width:100%;}
th {font-size:14pt;background-color:yellow;}
</style>

<title>Report zum Deaktivierungsscript</title>"

$datum = Get-Date -Format "dd.MM.yyyy, HH:mm"

$kombiniert = $Alt_Inaktiv__Deaktiviert + $Neu_Inaktiv_Deaktiviert

$body = $kombiniert | sort-object OU,AD_LastLogonDate | select SamAccountName,userprincipalname,Reason,AD_logonCount,OU,AD_LastLogonDate,AD_Logon_Age,CloudLogin,Cloud_LastLogonDate,Cloud_LogOnAge,Deaktivierung | ConvertTo-Html -Head $head -PreContent "<h1>Report erzeugt von $env:USERNAME am $datum</h1>"

##########################################################
$message = new-object System.Net.Mail.MailMessage
$message.From = $FROM
#$TO | %{ $message.To.Add($_) }
$message.To.Add($TO)
$message.IsBodyHtml = $True
$message.Subject = $Subject
$attach1 = new-object Net.Mail.Attachment($attachment1)
$message.Attachments.Add($attach1)
$attach2 = new-object Net.Mail.Attachment($attachment2)
$message.Attachments.Add($attach2)
$attach3 = new-object Net.Mail.Attachment($attachment3)
$message.Attachments.Add($attach3)
$attach4 = new-object Net.Mail.Attachment($attachment4)
$message.Attachments.Add($attach4)
$message.body = $body

$smtp = new-object Net.Mail.SmtpClient($SmtpServer, 25)
$smtp.EnableSsl = $false
$smtp.Credentials = $credential;
$smtp.Send($message)
########################################################## 