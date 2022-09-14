
$checkdata= @()

foreach ($m in $arbeitsliste) {

       $mail = ($m.mail).Trim()
         $AD = Get-ADUser -Filter {mail -eq $mail} -properties mail,userprincipalname,proxyaddresses,extensionAttribute4
        $Sam = $AD.Samaccountname
     $AP_alt = $AD.extensionAttribute4
     $AP_neu = ($m.AP).Trim()
 IF ($AP_neu -eq $AP_alt) { $AP_bleibt = $true}
        $UPN = $AD.userprincipalname
       $SMTP = "$($AD | select -expandproperty proxyaddresses | where { $_ -cmatch '^SMTP'} )" -replace "SMTP:"
     $target = "$($AD | select -expandproperty proxyaddresses | where { $_ -match "elka.mail.onmicrosoft.com$"} )" -replace "smtp:"
 $UPN_Domain = ($UPN -split "@")[1]
$SMTP_Domain = ($SMTP -split "@")[1]
  $UPN_alias = ($UPN -split "@")[0]
 $SMTP_alias = ($SMTP -split "@")[0]

$target = "$( $AD.proxyaddresses | where { $_ -match "mail.onmicrosoft.com$" } )" -replace "smtp:"

IF (!($target)) { $checkdata += "UPN $UPN has no proxy mail.onmicrosoft.com"           }
IF ($SMTP_domain -ne "elkw.de") { $checkdata += "Check $SMTP SMTP Domain $SMTP_domain"           }
IF ($UPN_domain -ne "elkw.de" ) { $checkdata += "Check UPN $UPN Domain $UPN_domain"              }
IF ($SMTP_alias -ne $UPN_alias) { $checkdata += "UPN $UPN_alias/SMTP $SMTP_alias Alias Mismatch" }
IF ($UPN_domain -match "elkw.local$" ) { $checkdata += "UPN $UPN has ELKW.LOCAL domain"           }
IF ($AP_neu -eq $AP_alt) { $AP_bleibt = $true ; $APcol = "white" } ELSE { $AP_bleibt = $false ; $APcol = "cyan" }

Write-host "SMTP: $SMTP | UPN: $UPN | Mail: $mail | " -F white -Nonewline

Write-host "$AP_neu `n" -F $APcol

try { Set-ADUser -Identity $Sam -Replace @{extensionattribute4=$AP_neu } -EA stop } catch { Write-host $Error[0].exception.message -F yellow }

$AP_check = (Get-ADUser -Filter {mail -eq $mail} -Properties extensionAttribute4).extensionAttribute4

Write-host "$AP_alt / $AP_neu || " -F $APcol -nonewline

IF ($AP_check -eq $AP_neu) { Write-host "AP Update Check OK:  $AP_check Success`n" -F green } ELSE { Write-host "AP Update Check FAIL: $AP_neu / $AP_check`n"}

} 

$checkdata | export-csv c:\temp\log.csv -D ";" -E UTF8 -NTI
$checkdata | Out-File c:\temp\log.txt -Encoding UTF8 ; $checkdata |FL