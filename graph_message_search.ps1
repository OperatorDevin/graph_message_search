# Loop to retrieve access token
do {
    # Prompt the user to enter the tenant ID
    $tenantId = Read-Host "Input your tenant ID"

    # Set the token endpoint and client credentials
    $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Prompt the user to enter the application ID
    $appId = Read-Host "Input your application ID"

    # Prompt the user to enter the client secret
    $clientSecret = Read-Host "Input your client secret"

    # Set the token request parameters
    $tokenParams = @{
        "grant_type" = "client_credentials"
        "client_id" = $appId
        "client_secret" = $clientSecret
        "scope" = "https://graph.microsoft.com/.default"
    }

    # Send the token request and retrieve the access token
    Write-Output "Sending token request..."
    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $tokenParams
        $accessToken = $tokenResponse.access_token
    }
    catch {
        Write-Host "Error obtaining access token. Please check the tenant ID, application ID, and client secret and try again."
        continue
    }

    # Check if access token was retrieved successfully
    if ($accessToken) {
        Write-Host "Access token retrieved successfully."
        break
    }
    else {
        Write-Output "Error obtaining access token. Please check the tenant ID, application ID, and client secret and try again."
    }
} while ($true)

# Loop to search for messages
do {
    # Prompt the user to enter the email address of the user whose mailbox they want to access
    $userEmail = Read-Host "Enter the email address of the user whose mailbox you want to access"

    # Prompt the user to enter the message ID
    $messageId = Read-Host "Enter the message ID"

    # Set the Graph API endpoint and message URL
    $graphEndpoint = "https://graph.microsoft.com/v1.0"
    $messageUrl = "$graphEndpoint/users/$userEmail/messages/$messageId"

    # Set the Graph API request headers
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }

    # Send the Graph API request to retrieve the message properties
    Write-Host "Retrieving message properties..."
    try {
        $messageProperties = Invoke-RestMethod -Uri $messageUrl -Headers $headers -Method Get
    }
    catch {
        Write-Host "Error retrieving message properties. Please check the email address and message ID and try again."
        continue
    }

    # Check if message properties were retrieved successfully
    if ($messageProperties) {
        # Output the message properties to the console
        $messageProperties | Select-Object subject, @{Name="Sender";Expression={$_.Sender.EmailAddress.Address}}, @{Name="SenderName";Expression={$_.Sender.EmailAddress.Name}}, @{Name="ToRecipients";Expression={$_.ToRecipients.EmailAddress.Address}}, @{Name="CcRecipients";Expression={$_.CcRecipients.EmailAddress.Address}}, @{Name="Body";Expression={$_.Body.Content}}

        # Prompt the user to save the output to a file
        $saveToFile = Read-Host "Do you want to save the message to a file? (Y/N)"
        if ($saveToFile -eq "Y") {
            # Prompt the user to enter the file path
            $filePath = Read-Host "Enter the file path to save the message (including the file extension)"
            $fileExtension = [System.IO.Path]::GetExtension($filePath)

            # Save the output to the specified file format
            switch ($fileExtension) {
                ".csv" {
                    $messageProperties | Export-Csv -Path $filePath -NoTypeInformation
                }
                ".xml" {
                    $messageProperties | Export-Clixml -Path $filePath
                }
                default {
                    Write-Host "Unsupported file format. Message not saved."
                }
            }

            # Output the file path to the console
            Write-Host "Message saved to file: $filePath"
        }
    }
    else {
        Write-Host "Error retrieving message properties. Please check the email address and message ID and try again."
    }

    # Prompt the user to search for another message or quit
    $searchAgain = Read-Host "Do you want to search for another message? (Y/N)"
} while ($searchAgain -eq "Y")

Write-Host "Exiting script..."
