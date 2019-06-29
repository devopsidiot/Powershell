$email = 'email'
$password = 'password'
$pass = ConvertTo-SecureString -AsPlainText $password -Force

$SecureString = $pass
$MySecureCreds = New-Object -TypeName System.Management.Automation.PSCredential $email,$SecureString

Send-MailMessage -To 'helpdesk@somecompany.com' -From 'some.admin@somecompany.com' -Subject "Daily Checklist - a person" -Body 'Do the thing.'-Credential $MySecureCreds -Attachments "C:\Users\person\Downloads\HChecklist.pdf" -UseSsl -SmtpServer 'smtp.office365.com'
