



$allgroups = get-ADGroup -Filter {organizationalunit -eq "OU=Test Accounts,DC=pshell,DC=site" }



$OUs = Get-OrganizationalUnit 

$DGS = $allgroups | where { $_.Groupscope -eq "Universal" -or $_.Groupscope -eq "Security" }

$OU_Select = $DGS.organizationalunit

 |out-gridview -T "select Group OU" -P |ft


$Domain_before = "pshell.site"

$Domain_after = "exo.red"

$groups = get-distributiongroup -ResultSize unlimited -organizationalunit "OU=Groups,OU=Test Accounts,DC=pshell,DC=site"

$groups |ft primarysmtpaddress,EmailAddresses,windowsemailaddress

foreach ($group in $groups) {

Set-DistributionGroup $group.distinguishedname -EmailAddressPolicyEnabled:$false

$SMTP =$group.PrimarySmtpAddress

$NEW_SMTP = $SMTP -replace ( $Domain_before, $Domain_after )

Set-DistributionGroup $group.distinguishedname -PrimarySmtpAddress $NEW_SMTP -Confirm:$false -force

Set-DistributionGroup $group.distinguishedname -WindowsEmailAddress $NEW_SMTP -Confirm:$false -force
<#
$primarysmtpaddress = $group.proxyaddresses | where { $_ -cmatch "^SMTP:" }
$Local_SMTP = $primarysmtpaddress | where { $_ -match "$Domain_before$"  }
$SMTP_After = $Local_SMTP -Replace ( $Domain_before, $Domain_after )
set-ADGroup $group.DistinguishedName |Set-ADObject -Replace @{ProxyAddresses="$SMTP_After"} 
#>
}

Get-AcceptedDomain
###########################################
$Domain_before = "pshell.site"

$Domain_after = "exo.red"

$ADgroups = Get-ADGroup -SearchBase "OU=Groups,OU=Test Accounts,DC=pshell,DC=site" -Filter * -Properties ProxyAddresses,DistinguishedName
$UniversalGroups = $ADgroups | where { $_.GroupScope -eq "Universal"}

foreach($group in $ADgroups) {

$primarysmtpaddress = $group.proxyaddresses | where { $_ -cmatch "^SMTP:" }
$Local_SMTP = $primarysmtpaddress | where { $_ -match "$Domain_before$"  }
$SMTP_After = $Local_SMTP -Replace ( $Domain_before, $Domain_after ) -replace "smtp:"
set-ADGroup $group.DistinguishedName |Set-ADObject -Replace @{ProxyAddresses="$SMTP_After"} 

}