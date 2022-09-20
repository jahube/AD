##########################################################################
$TenantId = "******************************************"
$clientSecret = "******************************************" 
$clientId = "******************************************"
##########################################################################
$tokenBody = @{  
    Grant_Type    = "client_credentials"  
    Scope         = "https://graph.microsoft.com/.default"  
    Client_Id     = $clientId  
    Client_Secret = $clientSecret  
}   

#$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenantid/oauth2/v2.0/token" -Method POST -Body $tokenBody  
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenantid/oauth2/v2.0/token" -Method POST -Body $tokenBody  -Proxy "http://proxy-dc.datagroup.local:3128" -ProxyUseDefaultCredentials

##########################################################################
$headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }
##########################################################################
$lastlogindates = @()

$Url = "https://graph.microsoft.com/beta/users?`$select=displayName,userPrincipalName,signInActivity,userType,assignedLicenses"
##########################################################################
While ($url -ne $Null) {
    $data = (Invoke-WebRequest -Headers $headers -Uri $url -Proxy "http://proxy-dc.datagroup.local:3128" -ProxyUseDefaultCredentials ) | ConvertFrom-Json
     #$data = (Invoke-WebRequest -Headers $headers -Uri $url) | ConvertFrom-Json
    $lastlogindates += $data.Value | select displayName,userPrincipalName,userType,Mail,  @{ N= "lastSignInDateTime" ; E = { $_.signInActivity.lastSignInDateTime } },@{ N= "lastNonInteractiveSignInDateTime" ; E = { $_.signInActivity.lastNonInteractiveSignInDateTime } }
    $url = $data.'@Odata.NextLink'
}
##########################################################################
$path = "C:\Scripte\Disable_AD_Users\Data"

$TS = (get-date -Format yyyy-MM-dd_HH.mm).ToString()

$lastlogindates.Where({$_.Usertype -ne "Guest"}) | Export-Csv "$path\Export-LastLoginDates-Cloud_$TS.csv" -Delimiter ";" -Encoding UTF8 -NTI -force
##########################################################################
