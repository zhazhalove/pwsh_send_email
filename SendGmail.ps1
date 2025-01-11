# Load Libraries
Add-Type -Path "$PSScriptRoot\refs\MailKit.dll"
Add-Type -Path "$PSScriptRoot\refs\MimeKit.dll"

# Create Email Message
$Message = New-Object MimeKit.MimeMessage
$Message.From.Add([MimeKit.MailboxAddress]::new("Sender Name", "user@gmail.com"))
$Message.To.Add([MimeKit.MailboxAddress]::new("Recipient Name", "user@example.com"))
$Message.Subject = 'Test Message'

# Email Body
$TextPart = [MimeKit.TextPart]::new("plain")
$TextPart.Text = "This is a test email sent using PowerShell with Gmail."
$Message.Body = $TextPart

# SMTP Client
$SMTP = New-Object MailKit.Net.Smtp.SmtpClient
try {
    # Connect to Gmail SMTP Server
    $SMTP.Connect('smtp.gmail.com', 587, [MailKit.Security.SecureSocketOptions]::StartTls)

    # Authenticate with Gmail SMTP Server
    $SMTP.Authenticate('user@gmail.com', 'app_password')

    # Send Email
    $SMTP.Send($Message)
    Write-Host "Email sent successfully!"
} catch {
    Write-Host "An error occurred: $_"
} finally {
    # Cleanup
    $SMTP.Disconnect($true)
    $SMTP.Dispose()
}
