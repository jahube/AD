
$UPN_report = @()

$routingdomain = "elkw.mail.onmicrosoft.com"

$path = "c:\temp"

#if (!(Test-Path $path)) { mkdir $path }

foreach ($mbx in $invalid_upn) {

$target = ""
#$alias = ($($mbx.primarysmtpaddress) -split '@')[0] -replace "smtp:"

$Ualias = ($($mbx.userprincipalname) -split '@')[0] -replace "smtp:"

$Domain = $($mbx.userprincipalname) -replace "$Ualias@"

$UPN_neu = $Ualias + "@" + "elkw.de"

#$target = $mbx.TargetAddress | where { $_ -match $routingdomain }

IF ($Domain -ne "elkw.de" -and $Domain -ne "elkw.local") { Write-host "$Ualias@" -ForegroundColor Green -NoNewline ; Write-host $domain -ForegroundColor Yellow }

IF ($Domain -eq "elkw.de" )
{ 

        Write-host "$($mbx.primarysmtpaddress) | $($mbx.primarysmtpaddress) `n" -ForegroundColor Green

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
    #   $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
    #   $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishedname
    #   $item | Add-Member -MemberType NoteProperty -Name GUID -Value $mbx.GUID
        $item | Add-Member -MemberType NoteProperty -Name msExchMailboxGuid -Value $mbx.msExchMailboxGuid
        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value "korrekt"
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value "korrekt"
        $item | Add-Member -MemberType NoteProperty -Name Error -Value "kein"

        $UPN_report += $item ;

}

IF ($Domain -eq "elkw.local" )
{

        #$target_neu = "$($alias + '@' + $routingdomain)"

        Write-host "$($mbx.userprincipalname)" -ForegroundColor cyan -NoNewline
        
        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
      # $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
      # $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishdedname
      # $item | Add-Member -MemberType NoteProperty -Name ExchangeObjectGUID -Value $mbx.ExchangeObjectGUID
        $item | Add-Member -MemberType NoteProperty -Name msExchMailboxGuid -Value $mbx.msExchMailboxGuid

   Try {

 Set-ADUser $mbx.samaccountname -Replace @{userprincipalname="$UPN_neu"} -ErrorAction Stop

 Write-host " | $UPN_neu erfolgreich gesetzt`n" -ForegroundColor green

Start-Sleep 1

        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value "wrong"
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value "successful"
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value $UPN_neu
        $item | Add-Member -MemberType NoteProperty -Name Error -Value "kein"
        
        $UPN_report += $item ;

       }
       
     catch

       {

 Write-host " | $UPN_neu setzen fehlgeschlagen `n" -ForegroundColor yellow
  Write-host $Error[0].Exception.Message -ForegroundColor yellow


          
        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value "wrong"
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value $UPN_neu
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value "failed"
        $item | Add-Member -MemberType NoteProperty -Name Error -Value $Error[0].Exception.Message

        $UPN_report += $item ;

       }
}



}


$path = "C:\Temp"

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$filepath = $path + '\UPN_report_' + $datestamp + '.CSV'

$UPN_report | Export-Csv -Path $filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force