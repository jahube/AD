#Publicfoldermbx
$batches = Get-MigrationBatch | where { ($_.identity -match "bezirk1" -or $_.identity -match "bezirk2") }   
$batches

$batches = Get-MigrationBatch | where { $_.identity -match "bezirk"}   
$batches

$MigrationUser = $batches | % { Get-MigrationUser -BatchId $_.batchguid } 

$abfrage = $MigrationUser | % { get-mailbox $_.MailboxEmailAddress  }

$abfrage  | where { ($_.defaultpublicfoldermailbox -ne "PFEND") } | ft userp*,prim*, recipienttypedetails,defaultpublicf*
$abfrage  | where { ($_.defaultpublicfoldermailbox -ne "PFEND") } | % { set-mailbox $_.Primarysmtpaddress -defaultpublicfoldermailbox PFEND  } 
$abfrage  | where { ($_.defaultpublicfoldermailbox -ne "PFEND") } | % { get-mailbox $_.Primarysmtpaddress  }  | ft userp*,prim*, recipienttypedetails,defaultpublicf*

$complete = $MigrationUser | ? { $_.status -eq "completed" }

$complete | % { set-mailbox $_.MailboxEmailAddress -defaultpublicfoldermailbox PFEND  } 

#abfrage 
$abfrage = $complete | % { get-mailbox $_.MailboxEmailAddress  }
$abfrage | ft userp*,prim*, recipienttypedetails,defaultpublicf*

########################################################################

#Shared doings
$FPbatches = Get-MigrationBatch | where { ($_.identity -match "Bezirk") -and $_.identity -match "funktion"}   
$FPbatches

$FPUser = $FPbatches | % { Get-MigrationUser -BatchId $_.batchguid } 
$FPcomplete = $FPuser | ? { $_.status -eq "completed" } 

$FPcomplete | % { set-mailbox $_.MailboxEmailAddress -type shared } 
$FPcomplete | % { set-mailbox $_.MailboxEmailAddress -MessageCopyforSentAsEnabled $true -MessageCopyforSendOnBehalfEnabled $true }

#abfrage shared
$abfrageFP = $FPcomplete | % { get-mailbox $_.MailboxEmailAddress  }
$abfrageFP | ft userp*,prim*, recipienttypedetails,*copy*