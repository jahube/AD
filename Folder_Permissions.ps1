
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$mbxs = get-mailbox -ResultSize unlimited

$folderstats =foreach ($mbx in $mbxs) { Get-MailboxFolderStatistics $mbx.DistinguishedName |select Identity,name,folderpath,foldertype,containerclass,ContentMailboxGuid,folderid }

$permissions = foreach ($folder in $folderstats) { Get-MailboxFolderPermission "$(($folder.identity -split "\\")[0]):$($folder.FolderId)" | Where { $_.user.usertype -ne "Default" -and $_.user.usertype -ne "anonymous" } | select @{n = "Name"; E={ $_.identity.MailBoxOwnerId.Name }},foldername,@{n = "FolderID"; E={ $_.identity.StoreObjectId }},@{n = "ObjectGuid"; E={ $_.identity.MailBoxOwnerId.ObjectGuid.Guid }},@{n = "IsDeleted"; E={ $_.identity.MailBoxOwnerId.IsDeleted }},@{n = "OU"; E={ $_.identity.MailBoxOwnerId.Parent }},@{n = "DistinguishedName"; E={ $_.identity.MailBoxOwnerId.DistinguishedName }},@{n = "DN"; E={ $_.identity.MailBoxOwnerId.Rdn }},@{n = "accessrights"; E={ $_.accessrights }},@{n = "Usertype"; E={ $_.user.usertype }},@{n = "Displayname"; E={ $_.user.Displayname }},@{n = "ADRecipient"; E={ $_.user.ADRecipient }}}


# ($folderstats[20].Identity -split "\\")[0]

$permissions  = foreach ($folder in $folderstats) { Get-MailboxFolderPermission "$(($folder.identity -split "\\")[0]):$($folder.FolderId)" | Where { $_.user.usertype -ne "Default" -and $_.user.usertype -ne "anonymous" } | select @{n = "Name"; E={ $_.identity.MailBoxOwnerId.Name }},foldername,`
@{n = "FolderID"; E={ $_.identity.StoreObjectId }},@{n = "ObjectGuid"; E={ $_.identity.MailBoxOwnerId.ObjectGuid.Guid }},@{n = "IsDeleted"; E={ $_.identity.MailBoxOwnerId.IsDeleted }},`
@{n = "OU"; E={ $_.identity.MailBoxOwnerId.Parent }},@{n = "DistinguishedName"; E={ $_.identity.MailBoxOwnerId.DistinguishedName }},@{n = "DN"; E={ $_.identity.MailBoxOwnerId.Rdn }},@{n = "accessrights"; E={ $_.accessrights }},`
@{n = "Usertype"; E={ $_.user.usertype }},@{n = "Displayname"; E={ $_.user.Displayname }},@{n = "ADRecipient"; E={ $_.user.ADRecipient }}}

$permissions | Export-csv C:\temp\Folderpermissions_all.csv -Delimiter ";" -Encoding UTF8 -Append

#| Export-csv C:\temp\Folderpermissions2.csv -Delimiter ";" -Encoding UTF8 -Append

foreach ($folder in $folderstats) { Get-MailboxFolderPermission "$(($folder.identity -split "\\")[0]):$($folder.FolderId)" | Where { $_.user.usertype -ne "Default" -and $_.user.usertype -ne "anonymous" } | select @{n = "Name"; E={ $_.identity.MailBoxOwnerId.Name }},foldername,`
@{n = "FolderID"; E={ $_.identity.StoreObjectId }},@{n = "ObjectGuid"; E={ $_.identity.MailBoxOwnerId.ObjectGuid.Guid }},@{n = "IsDeleted"; E={ $_.identity.MailBoxOwnerId.IsDeleted }},`
@{n = "OU"; E={ $_.identity.MailBoxOwnerId.Parent }},@{n = "DistinguishedName"; E={ $_.identity.MailBoxOwnerId.DistinguishedName }},@{n = "DN"; E={ $_.identity.MailBoxOwnerId.Rdn }},@{n = "accessrights"; E={ $_.accessrights }},`
@{n = "Usertype"; E={ $_.user.usertype }},@{n = "Displayname"; E={ $_.user.Displayname }},@{n = "ADRecipient"; E={ $_.user.ADRecipient }} | Export-csv C:\temp\Folderpermissions_append.csv -Delimiter ";" -Encoding UTF8 -Append }


# Where { $_.user.usertype -ne "Default" -and $_.user.usertype -ne "anonymous" }


# http://aka.ms/personaltagscript

# http://aka.ms/createlabusers
