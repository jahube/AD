
$DOMAINDATA = @()

$skip_admin = $true

$routingdomain = "cvgdnsabr.mail.onmicrosoft.com"

$Domain_before = "pshell.site"

$Domain_after = "exo.red"

$mbxs = get-mailbox -ResultSize unlimited

$path = "c:\temp"

foreach ($mbx in $mbxs) {

 $Local_UPN = "" ; 

$Local_SMTP = "" ;

     $alias = ($($mbx.primarysmtpaddress) -split '@')[0]

 IF (($alias -match "administrator" -or $alias -match "^adm_") -and $skip_admin) { $not_admin = $false } else { $not_admin = $true }

 $Local_UPN = $mbx.userprincipalname | where { $_ -match $Domain_before }

$Local_SMTP = $mbx.primarysmtpaddress | where { $_ -match $Domain_before }

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name alias -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name distinguishedname -Value $mbx.distinguishedname
        $item | Add-Member -MemberType NoteProperty -Name GUID -Value $mbx.GUID
        $item | Add-Member -MemberType NoteProperty -Name ExchangeGUID -Value $mbx.ExchangeGUID


IF (!($Local_UPN) -and !($Local_SMTP))

{
        Write-host "SMTP: $($mbx.userprincipalname) | UPN: $($mbx.userprincipalname) - Korrekt`n`n" -ForegroundColor Green

        $item | Add-Member -MemberType NoteProperty -Name existing_SMTP -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name missing_SMTP -Value "SMTP korrekt"
        $item | Add-Member -MemberType NoteProperty -Name SMTP_added -Value "SMTP korrekt"
        $item | Add-Member -MemberType NoteProperty -Name SMTP_Error -Value "kein"

        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value "UPN updated"
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value "UPN korrekt"
        $item | Add-Member -MemberType NoteProperty -Name UPN_Error -Value "kein"

        $DOMAINDATA += $item ;
}

 # primarysmtpaddress

 IF ($Local_SMTP -and $not_admin)

 {
        $SMTP_After = $Local_SMTP -replace ( $Domain_before, $Domain_after )

         Write-host "SMTP: before $($mbx.primarysmtpaddress) | " -ForegroundColor cyan -NoNewline

        Try {

 Set-mailbox -Identity $mbx.userprincipalname -primarysmtpaddress $SMTP_After -CF:$false -ErrorAction stop ;

 Write-host "SMTP: after $SMTP_After - SUCCESS`n" -ForegroundColor green

        $item | Add-Member -MemberType NoteProperty -Name existing_SMTP -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name missing_SMTP -Value "SMTP updated"
        $item | Add-Member -MemberType NoteProperty -Name SMTP_added -Value $SMTP_After
        $item | Add-Member -MemberType NoteProperty -Name SMTP_Error -Value "kein"

            }

          catch

            {

 Write-host "SMTP: after $($mbx.primarysmtpaddress) - SMTP $SMTP_After UPDATE FAILED" -ForegroundColor red -BackgroundColor Yellow

        $item | Add-Member -MemberType NoteProperty -Name existing_SMTP -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name missing_SMTP -Value $SMTP_After
        $item | Add-Member -MemberType NoteProperty -Name SMTP_added -Value "SMTP update Failed"
        $item | Add-Member -MemberType NoteProperty -Name SMTP_Error -Value $Error[0].Exception.Message

            }

}

 # userprincipalname

 IF ($Local_UPN -and $not_admin)

 {
        Write-host "UPN: before $($mbx.userprincipalname) | " -ForegroundColor cyan -NoNewline

        $UPN_After = $Local_UPN -replace ( $Domain_before, $Domain_after )

        Try {

 Set-mailbox -Identity $mbx.userprincipalname -userprincipalname $UPN_After -CF:$false -ErrorAction stop ;

 Write-host "UPN: after $UPN_After - SUCCESS`n`n" -ForegroundColor green

        # userprincipalname

        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value "UPN updated"
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value $UPN_After
        $item | Add-Member -MemberType NoteProperty -Name UPN_Error -Value "kein"

            }

          catch

            {
            
 Write-host "UPN: after $($mbx.userprincipalname) - UPN $UPN_After UPDATE FAILED`n`n" -ForegroundColor red -BackgroundColor Yellow

        $item | Add-Member -MemberType NoteProperty -Name existing_UPN -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name missing_UPN -Value $UPN_After
        $item | Add-Member -MemberType NoteProperty -Name UPN_added -Value "UPN update Failed"
        $item | Add-Member -MemberType NoteProperty -Name UPN_Error -Value $Error[0].Exception.Message

            }

}

      $DOMAINDATA += $item ;

}

# Export CSV Data

if (!(Test-Path $path)) { mkdir $path }

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM

$filepath = $path + '\DOMAIN_Update_DATA_' + $datestamp + '.CSV'

$DOMAINDATA | Export-Csv -Path $filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force

# END
