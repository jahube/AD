 $outbound = Get-SendConnector | where { $_.identity -match "^outbound" } 

 $outbound | % { $_ | get-ADPermission | where { $_.extendedrights -match "ms-exch-send-headers-routing" } | ft -autosize }

 # https://www.alitajran.com/remove-message-header-in-exchange-server/