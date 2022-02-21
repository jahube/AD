
$target_report = @()

$routingdomain = "cvgdnsabr.mail.onmicrosoft.com"

$mbxs = get-mailbox -ResultSize unlimited

$path = "c:\temp"

if (!(Test-Path $path)) { mkdir $path }

foreach ($mbx in $mbxs) {

$target = ""
$alias = ($($mbx.primarysmtpaddress) -split '@')[0]

$target = $mbx.emailaddresses | where { $_ -match $routingdomain }

IF ($target)  

{

        Write-host "$($mbx.primarysmtpaddress) | $target `n" -ForegroundColor Green

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishedname
        $item | Add-Member -MemberType NoteProperty -Name GUID -Value $mbx.GUID
        $item | Add-Member -MemberType NoteProperty -Name ExchangeGUID -Value $mbx.ExchangeGUID
        $item | Add-Member -MemberType NoteProperty -Name existing_target -Value $target.SmtpAddress
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
        $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name distinguishdedname -Value $mbx.distinguishdedname
        $item | Add-Member -MemberType NoteProperty -Name ExchangeObjectGUID -Value $mbx.ExchangeObjectGUID
        $item | Add-Member -MemberType NoteProperty -Name ExchangeGUID -Value $mbx.ExchangeGUID

        $target_report += $item ;

        Try {

 Set-mailbox -Identity $mbx.userprincipalname -EmailAddresses @{Add="$target_neu"} -ErrorAction stop ;

 Write-host " | $target_neu erfolgreich gesetzt`n" -ForegroundColor green

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

}

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$filepath = $path + '\target_report_' + $datestamp + '.CSV'

$target_report | Export-Csv -Path $filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force