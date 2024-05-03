# Create an Outlook.Application object
$outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$namespace = $outlook.GetNamespace("MAPI")

# Get the Inbox folder
$inbox = $namespace.GetDefaultFolder(6)

# Get the first six items in the Inbox
$emails = $inbox.Items | Select-Object -First 22

# Loop through each email and print details
foreach ($email in $emails) {
    Write-Host "Subject: $($email.Subject)"
    Write-Host "Received: $($email.ReceivedTime)"
    Write-Host "Sender: $($email.SenderName)"
    Write-Host "Body: $($email.Body)"
    Write-Host ("-" * 80)
}