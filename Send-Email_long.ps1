Function Send-EMail {
    Param (
        [Parameter(`
            Mandatory=$true)]
        [String]$To,
         [Parameter(`
            Mandatory=$true)]
        [String]$From,
        [Parameter(`
            mandatory=$true)]
        [String]$Password,
        [Parameter(`
            Mandatory=$true)]
        [String]$Subject,
        [Parameter(`
            Mandatory=$true)]
        [String]$Body,
        [Parameter(`
            Mandatory=$true)]
        [String]$attachment
    )
    try
        {
            $Msg = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
            $Srv ="smtp.mail.me.com"#"smtp.gmail.com" # "127.0.0.1" #  "smtp.protonmail.com" #mail.protonmail.com
            if ($attachment -ne $null) {
                try
                    {
                        $Attachments = $attachment -split (",");
                        ForEach ($val in $Attachments)
                            {
                                $attch = New-Object System.Net.Mail.Attachment($val)
                                $Msg.Attachments.Add($attch)
                            }
                    }
                catch
                    {
                        exit 2; 
                    }
            }
            $Client = New-Object Net.Mail.SmtpClient($Srv, 587) #587 port for smtp.gmail.com SSL, 1025 for proton
            $Client.EnableSsl = $True 
            $Client.Credentials = New-Object System.Net.NetworkCredential($From, $Password); #.Split("@")[0]
            $Client.Send($Msg)
            Remove-Variable -Name Client
            Remove-Variable -Name Password
            exit 7; 
          }
      catch
          {            $_
            exit 3;   
          }
} #End Function Send-EMail
try
    {        $Cred= Get-Credential -Message "enter Your email credentials"            Send-EMail -attachment "C:\Crash Course in Azure DevOps for Project Management.docx" ` #,C:\Users\alcanto\OneDrive - Microsoft\Bot Capabilities.xlsx"`  
        -To "$Cred.UserName" -Body "Hello" -Subject "test" -password $Cred.Password -From "$Cred.UserName"
    }
catch
    {
        exit 4; 
    }