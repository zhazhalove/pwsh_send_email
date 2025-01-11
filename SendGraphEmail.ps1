<#
    Registered application
      - Graph API Application
        - API Permissions
          - Mail.Send
      - Authentication
        - Client secrets
#>
$TenantId = ""
$ApplicationId = ""
$ClientSecret = ""
$SenderEmail = "user@tenant" # MUST have an active Exchange Online mailbox
$RecipientEmail = "user@example.com"
$EmailSubject = "Hello from PowerShell Graph API"
$EmailBodyContent = "This is a test email sent using Microsoft Graph API."

try {
    # Convert the client secret to a secure string
    $Pass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force

    # Create a credential object using the client ID and secure string
    $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $Pass

    # Connect to Microsoft Graph with Client Secret
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Cred

    $params = @{
        message = @{
            subject = $EmailSubject
            body = @{
                contentType = "Text"
                content = $EmailBodyContent
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $RecipientEmail
                    }
                }
            )
            # ccRecipients = @(
            # 	@{
            # 		emailAddress = @{
            # 			address = "danas@contoso.com"
            # 		}
            # 	}
            # )
        }
        saveToSentItems = "true"
    }

    Send-MgUserMail -UserId $SenderEmail -BodyParameter $params

}
catch {
    Write-Host "Encountered Error: $_.Exception.Message()"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
