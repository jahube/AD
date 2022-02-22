$USER = "affected@user"

$OfflineMode = $false

# desktop/MS-Logs+Timestamp

$ts = Get-Date -Format yyyyMMdd_hhmmss

$DesktopPath = "C:\temp"

$logsPATH = mkdir "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$ts"

Start-Transcript "$logsPATH\Transcript_$ts.txt"

$FormatEnumerationLimit = -1
[int]$start = "-90"

$data = Search-MailboxAuditLog -Identity $user -ShowDetails -StartDate (get-date).AddDays($start) -EndDate (get-date)

$data | Export-Clixml "$logsPATH\mailboxlogs.xml"

get-mailbox $user | select AuditEnabled
get-mailbox $user | select -expandproperty auditadmin
get-mailbox $user | select -expandproperty auditdelegate
get-mailbox $user | select -expandproperty auditowner

# Open File Dialog - Offline mode
$data | select-object operation,clientprocessname,clientinfostring,ClientVersion,clientip -Unique | ft > "$logsPATH\Types-Unique.txt"

# deletion actions only
$softdelete = $data | where {$_.operation -eq "softdelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$harddelete = $data | where {$_.operation -eq "harddelete" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList
$MoveToDltd = $data | where {$_.operation -eq "MoveToDeletedItems" } | select operation,clientprocessname,clientinfostring,lastaccessed,clientip,ClientVersion,FolderPathName,SourceItemFolderPathNamesList,SourceItemSubjectsList

# all actions
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip | select count,Name | Sort count -Descending | ft > "$logsPATH\Types-Summary.txt"

$harddelete | FT > "$logsPATH\harddelete-Details.txt"
$harddelete | Export-CSV "$logsPATH\harddelete-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
$harddelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\harddelete.txt"
$harddelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\harddelete-Summary.txt"

$softdelete | FT > "$logsPATH\softdelete-Details.txt"
$softdelete | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\softdelete.txt"
$softdelete | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\softdelete-Summary.txt"
$softdelete | Export-CSV "$logsPATH\softdelete-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$MoveToDltd | FT > "$logsPATH\MoveToDltd-Details.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,ClientVersion,clientip,LogonType,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\MoveToDltd.txt"
$MoveToDltd | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion | select count,Name | Sort count,Operation -Descending > "$logsPATH\MoveToDltd-Summary.txt"
$MoveToDltd | Export-CSV "$logsPATH\MoveToDltd-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

$data | FT > "$logsPATH\ALL-Details.txt"
$data | group operation,clientprocessname,clientinfostring,ClientVersion,LogonType,clientip,CrossMailboxOperation,DestMailboxOwnerUPN,ExternalAccess | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL.txt"
$data | group operation,clientprocessname,clientinfostring,LogonType,ClientVersion,clientip | select count,Name | Sort count,Operation -Descending > "$logsPATH\ALL-Summary.txt"
$data | Export-CSV "$logsPATH\ALL-Details.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force

Compress-Archive -Path $logsPATH -DestinationPath "$DesktopPath\MS-Logs\Mailbox-Audit-Logs_$($USER.replace('@',"-"))_$ts.Zip" -Force # Zip Logs
Invoke-Item $DesktopPath\MS-Logs # open file manager