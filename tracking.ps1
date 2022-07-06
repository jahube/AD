$logs = get-transportservice | invoke-command { Get-MessageTrackingLog -Start (get-date).adddays(-5) -resultsize unlimited | where { $_.recipients -match "DATA.com" } };

$logs | select *,{$_.Recipients},{$_.recipientstatus},{ $_.EventData } -ExcludeProperty recipients,recipientstatus,EventData |export-csv c:\Temp\data.csv -delimiter ";" -Force



$logs2 = get-transportservice | invoke-command { Get-MessageTrackingLog -sender sender@domain.com -Start (get-date).adddays(-90) -resultsize unlimited }

$logs2 | select *,{$_.Recipients},{$_.recipientstatus},{ $_.EventData } -ExcludeProperty recipients,recipientstatus,EventData |export-csv c:\Temp\data2.csv -delimiter ";" -Force
