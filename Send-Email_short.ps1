$Cred = Get-Credential -Message "enter Your email credentials"
$address = $Cred.UserName
$archive = 
$Body = @"

Give a man a script; 
    feed him for a Get-Date. 
`r`n
Teach a man to script; 
    feed him for a New-TimeSpan.
`r`n 

                Lao Tzu, 4th century BC
`r`n 

"@

$Email = @{
From = $address
To = $address
Subject = "Email From Powershell"
Body = "$Body"
Credential = $Cred
SMTPServer = "smtp.mail.me.com"
Port = "587"
UseSsl = $true
}

Send-MailMessage @Email