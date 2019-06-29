Get-EventLog -LogName System -EntryType Error -Newest 500 | Export-Csv "C:\Scripts\$hostname EventLog_$(get-date -f MM-dd-yyyy).csv"

$Hostname = hostname
$OL = New-Object -comobject outlook.application

$mItem = $OL.CreateItem("olMailItem")

$mItem.To = "support@somecompany.com"
$mItem.Subject = "$Hostname EventLog"
$mItem.Body = "$Hostname EventLog_$(get-date -f MM-dd-yyyy)"
$mItem.Attachments.Add("C:\Scripts\EventLog_$(get-date -f MM-dd-yyyy).csv")
$mItem.Attachments.Add("C:\Windows\MEMORY.DMP")

$mItem.Send()
