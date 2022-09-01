
$Path = "C:\temp\Locked-out_users"

$DCs = (Get-ADDomainController -filter { name -like "DGDC-DC*"}).name

$lockedOut_Today = $DCs | % { invoke-command { Get-WinEvent -ea SilentlyContinue -ComputerName $_ -FilterHashtable @{ providername="Microsoft-Windows-Security-Auditing" ; LogName="Security"; ID=4740; StartTime = [datetime]::today } } }

$Data = @()
Foreach($Event in $lockedOut_Today )
  {

$Item =  $Event | Select-Object -Property @(
        @{Label = 'User'; Expression = {$_.Properties[0].Value}}
        @{Label = 'DomainController'; Expression = {$_.MachineName}}
        @{Label = 'EventId'; Expression = {$_.Id}}
        @{Label = 'LockoutTimeStamp'; Expression = {$_.TimeCreated}}
        @{Label = 'Message'; Expression = {$_.Message -split "`r" | Select -First 1}}
        @{Label = 'LockoutSource'; Expression = {$_.Properties[1].Value}}
      )
$Item
$Data += $Item

   }

$timestamp = get-date -Format yyyy-MM-dd_HH.mm.ss

$file = "LockedOut_Users-Erstellungsdatum-$timestamp.csv"
$Data | export-csv "$Path\$file" -Delimiter ";" -Encoding UTF8 -NTI -Force

#Menü Freischalten - Unlock-ADAccount
$Data | Out-GridView -PassThru -Title "SELECT USERS TO UNLOCK" | % { Unlock-ADAccount $_.User -Confirm:$false }

#$Data|select DomainController,eventid,Lockoutsource,lockouttimestamp,@{n="mail"; E= { (get-aduser $_.user -Properties mail).mail }},@{n="manager"; E= {(get-aduser (get-aduser $_.user -Properties manager).manager -Properties mail).mail }}

# Email + Manager
$Data| ft DomainController,eventid,Lockoutsource,lockouttimestamp,@{n="mail"; E= { (get-aduser $_.user -Properties mail).mail }},@{n="manager"; E= {(get-aduser (get-aduser $_.user -Properties manager).manager -Properties mail).mail }}