
$Success = @()
$Fail = @()

$Domain_before = "pshell.site"

$Domain_after = "exo.red"

$mbx = Get-Mailbox -ResultSize unlimited

foreach ($mbx in $mbxs) {

$Proxys = $mbx.emailaddresses -split "," | where { $_ -match "$Domain_before$" }

foreach ($Proxy in $Proxys) {

IF ($Proxy -notmatch "^SIP:") {

    Try
      
     {

Set-Mailbox $mbx.userprincipalname -emailaddresses @{Remove="$(@($Proxy.proxyaddressstring).ToString())"} -Confirm:$false -ErrorAction silentlycontinue

Write-host "SMTP: $($mbx.userprincipalname) | Proxyaddress: $("$Proxy") - Korrekt entfernt`n`n" -ForegroundColor Green

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name proxyaddress -Value $Proxy
        $item | Add-Member -MemberType NoteProperty -Name RESULT -Value "SUCCESS"
        $item | Add-Member -MemberType NoteProperty -Name emailaddresses -Value "$(($mbx.emailaddresses -split ",") -join ' | ')"

     $Success +=  $item

     } 
     
   Catch 
     
     {

     Write-host "SMTP: $($mbx.userprincipalname) | Proxyaddress: $("$Proxy") - FAILED`n`n" -ForegroundColor red -BackgroundColor Yellow

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name proxyaddress -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name RESULT -Value "FAIL"
        $item | Add-Member -MemberType NoteProperty -Name emailaddresses -Value "$(($mbx.emailaddresses -split ",") -join ' | ')"

        $fail +=  $item

     }

}

IF ($Proxy -match "^SIP:") {

$alias = ($($mbx.primarysmtpaddress) -split '@')[0]

$New_SIP = "SIP:" + $alias + "@" + $Domain_after

    Try
      
     {

Set-Mailbox $mbx.userprincipalname -emailaddresses @{Remove="$(@($Proxy.proxyaddressstring).ToString())"} -Confirm:$false -ErrorAction silentlycontinue

Set-Mailbox $mbx.userprincipalname -emailaddresses @{Add="$New_SIP"} -Confirm:$false -ErrorAction Stop

Write-host "SMTP: $($mbx.userprincipalname) | Proxyaddress: $("$Proxy") - Korrekt entfernt`n`n" -ForegroundColor Green

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name proxyaddress -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name RESULT -Value "SUCCESS"
        $item | Add-Member -MemberType NoteProperty -Name emailaddresses -Value "$(($mbx.emailaddresses -split ",") -join ' | ')"

        $Success +=  $item
     } 
     
   Catch 
     
     {

   Write-host "SMTP: $($mbx.userprincipalname) | Proxyaddress: $("$Proxy") - FAILED`n`n" -ForegroundColor red -BackgroundColor Yellow

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $mbx.userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $mbx.primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name proxyaddress -Value $mbx.alias
        $item | Add-Member -MemberType NoteProperty -Name RESULT -Value "FAIL"
        $item | Add-Member -MemberType NoteProperty -Name emailaddresses -Value "$(($mbx.emailaddresses -split ",") -join ' | ')"

        $fail +=  $item
     }

}

  }
}