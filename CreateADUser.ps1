Import-Module ActiveDirectory
$ADUsers = Import-CSV C:\Test\Newhire.csv
$Domain = "Domain.com"
$UseSearchbase = $false

$AlreadyExists = @()
 $CreateFailed = @()
$CreateSuccess = @()

Foreach ($user in $username)
{

$Param = @{   Enabled = $true
       SamAccountName = $user.username
    UserprincipalName = $($user.username + '@' + $domain)
                 Name = $($(@($user.firstname).Trim()) + ' ' + $(@($user.lastname).Trim()))
            GivenName = ($user.firstname)
              Surname = ($user.lastname)
ChangePasswordAtLogon = $false
          Displayname = $($(@($user.lastname).Trim()) + ', ' + $(@($user.firstname).Trim()))
           Department = $user.department
                 Path = $user.ou
      AccountPassword = (Convertto-SecureString $User.Password -AsPlainText -Force) }

       If ($UseSearchbase) 
        {
            $usercheck = $(try {Get-ADUser -SearchBase $($user.ou) -F {SamAccountName -eq $($User.username)} -EA stop } catch {$null})
        }
       Else
        {
            $usercheck = $(try {Get-ADUser -F {SamAccountName -eq $($User.username)} -EA stop } catch {$null})
        }

       If ($usercheck)
       {
           Write-Warning "A user account $($username.username) already exists in Active Directory."
           $AlreadyExists += [Array]$Param
       }
       else
       {
          Try { New-AdUser @Param -EA stop ; $CreateSuccess += [Array]$Param } Catch { $CreateFailed += [Array]$Param }
       }
}
