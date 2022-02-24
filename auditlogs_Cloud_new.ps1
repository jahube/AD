$USER = "USER@domain.de"

$OfflineMode = $false

# desktop/MS-Logs+Timestamp
$ts = Get-Date -Format yyyyMMdd_hhmmss

$DesktopPath = "C:\temp"

$logsPATH = mkdir "$DesktopPath\MS-Logs\schaeffer_datagroup_14_$ts"

Start-Transcript "$logsPATH\Transcript_$ts.txt"

$FormatEnumerationLimit = -1

[int]$start = "-21" # Tage zurück

$data = Search-MailboxAuditLog -Identity $user -ShowDetails -StartDate (get-date).AddDays($start) -EndDate (get-date)

$data | Export-Clixml "$logsPATH\mailboxlogs.xml"

get-mailbox $user | select AuditEnabled
get-mailbox $user | select -expandproperty auditadmin
get-mailbox $user | select -expandproperty auditdelegate
get-mailbox $user | select -expandproperty auditowner

$data | select-object operation,clientprocessname,clientinfostring,ClientVersion,clientip -Unique | ft > "$logsPATH\Types-Unique.txt"

# deletion actions only
$softdelete = $data | where {$_.operation -eq "softdelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$harddelete = $data | where {$_.operation -eq "harddelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$MoveToDltd = $data | where {$_.operation -eq "MoveToDeletedItems" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$Update = $data | where {$_.operation -eq "Update" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$Create = $data | where {$_.operation -eq "Create" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList

# all actions
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip | select count,Name | Sort count -Descending | ft > "$logsPATH\Types-Summary.txt"
$harddelete | group lastaccessed | sort-object lastaccessed |select name, count, lastaccessed | FT > "$logsPATH\harddelete-Details.txt"
$harddelete | Export-CSV "$logsPATH\harddelete-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
$harddelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort lastaccessed,count,Operation -Descending > "$logsPATH\harddelete.txt"
$harddelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\harddelete-Summary.txt"

$softdelete | FT > "$logsPATH\softdelete-Details.txt"
$softdelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort lastaccessed,count,Operation -Descending > "$logsPATH\softdelete.txt"
$softdelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\softdelete-Summary.txt"
$softdelete | Export-CSV "$logsPATH\softdelete-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$MoveToDltd | FT > "$logsPATH\MoveToDltd-Details.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort lastaccessed,count,Operation -Descending > "$logsPATH\MoveToDltd.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\MoveToDltd-Summary.txt"
$MoveToDltd | Export-CSV "$logsPATH\MoveToDltd-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$Update | FT > "$logsPATH\Update-Details.txt"
$Update | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort lastaccessed,count,Operation -Descending > "$logsPATH\Update.txt"
$Update | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\Update-Summary.txt"
$Update | Export-CSV "$logsPATH\Update-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$Create | FT > "$logsPATH\Create-Details.txt"
$Create | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort lastaccessed,count,Operation -Descending > "$logsPATH\Create.txt"
$Create | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\Create-Summary.txt"
$Create | Export-CSV "$logsPATH\Create-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$data | FT > "$logsPATH\ALL-Details.txt"
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL.txt"
$data | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion,clientip | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL-Summary.txt"
$data | Export-CSV "$logsPATH\ALL-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
Compress-Archive -Path $logsPATH -DestinationPath "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$($USER.replace('@',"-"))_$ts.Zip" -Force # Zip Logs
Invoke-Item $DesktopPath\MS-Logs # open file manager


$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip | select count,Name | Sort count -Descending | ft > "$logsPATH\Types-Summary.txt"
