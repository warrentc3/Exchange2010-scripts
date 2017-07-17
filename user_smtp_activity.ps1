##
##
Get-MessageTrackingLog -resultsize unlimited | where {$_.Source -like 'SMTP' -and (($_.Sender -like 'user@domain.com') -or ($_.Recipients -like '*user@domain.com*'))} | select Sender, @{l="Recipients";e={$_.Recipients -join " ;"}},MessageSubject, Timestamp | Export-Csv -NoTypeInformation c:\exportpath\user_smtp.csv
