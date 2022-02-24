
$target_report = @()

$routingdomain = "xxx.mail.onmicrosoft.com"

$path = "c:\temp"

if (!(Test-Path $path)) { mkdir $path }

foreach ($mbx in $MISSING_TARGET) {

$target = ""

$alias = ($($mbx.primarysmtpaddress) -split '@')[0] -replace "smtp:"

$Domain = $($mbx.primarysmtpaddress) -replace "$alias@"

$target = $mbx.TargetAddress | where { $_ -match $routingdomain }

IF ($Domain -eq "elkw.de" -or $Domain -match "sbv") { 

IF ($target)  

{

        Write-host "$($mbx.primarysmtpaddress) | $target `n" -ForegroundColor Green

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
    #   $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
    #   $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishedname
    #   $item | Add-Member -MemberType NoteProperty -Name GUID -Value $mbx.GUID
        $item | Add-Member -MemberType NoteProperty -Name msExchMailboxGuid -Value $mbx.msExchMailboxGuid
        $item | Add-Member -MemberType NoteProperty -Name existing_target -Value $target.TargetAddress
        $item | Add-Member -MemberType NoteProperty -Name missing_target -Value "vorhanden"
        $item | Add-Member -MemberType NoteProperty -Name target_added -Value "vorhanden"
        $item | Add-Member -MemberType NoteProperty -Name Error -Value "kein"

        $target_report += $item ;

              }

             ELSE

              {

        $target_neu = "$($alias + '@' + $routingdomain)"

        Write-host "$($mbx.primarysmtpaddress)" -ForegroundColor cyan -NoNewline
        
        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
      # $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
      # $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishdedname
      # $item | Add-Member -MemberType NoteProperty -Name ExchangeObjectGUID -Value $mbx.ExchangeObjectGUID
        $item | Add-Member -MemberType NoteProperty -Name msExchMailboxGuid -Value $mbx.msExchMailboxGuid
        
        Try {

 Set-ADUser $mbx.samaccountname -add @{ProxyAddresses="smtp:$target_neu"} -ErrorAction Stop

 #Set-mailbox -Identity $mbx.userprincipalname -EmailAddresses @{Add="$target_neu"} -ErrorAction stop ;

 Write-host " | $target_neu erfolgreich gesetzt`n" -ForegroundColor green

Start-Sleep 1

        $item | Add-Member -MemberType NoteProperty -Name existing_target -Value "missing"
        $item | Add-Member -MemberType NoteProperty -Name missing_target -Value "successful"
        $item | Add-Member -MemberType NoteProperty -Name target_added -Value $target_neu
        $item | Add-Member -MemberType NoteProperty -Name Error -Value "kein"
        
        $target_report += $item ;

            }

          catch

            {

 Write-host " | $target_neu setzen fehlgeschlagen `n" -ForegroundColor yellow
          
        $item | Add-Member -MemberType NoteProperty -Name existing_target -Value "missing"
        $item | Add-Member -MemberType NoteProperty -Name missing_target -Value $target_neu
        $item | Add-Member -MemberType NoteProperty -Name target_added -Value "failed"
        $item | Add-Member -MemberType NoteProperty -Name Error -Value $Error[0].Exception.Message

        $target_report += $item ;

            }
}
  } ELSE { Write-host "$alias@" -ForegroundColor Green -NoNewline ; Write-host $domain -ForegroundColor Yellow }

}


$path = "C:\Temp"

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$filepath = $path + '\target_report_' + $datestamp + '.CSV'

$target_report | Export-Csv -Path $filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force