# graph_message_search
Powershell Script to Search for Messages via Graph API

This PowerShell script allows you to search for and retrieve the properties of a specific email message in a user's mailbox using the Microsoft Graph API.

## Prerequisites

Before you can use this script, you must have the following:

- A Microsoft Azure account with an active subscription
- An Azure Active Directory (AD) tenant
- An Azure AD application with the following permissions:
  - Application permissions: Mail.Read
  - Delegated permissions: User.Read
- The application ID and client secret for the Azure AD application
- The tenant ID for your Azure AD tenant

## Usage

To use this script, follow these steps:

1. Open a PowerShell console or terminal window.
2. Navigate to the directory where you saved the script.
3. Run the script by entering the following command: `.\graph_message_search.ps1`
4. Follow the prompts to enter your Azure AD tenant ID, application ID, and client secret.
5. Enter the email address of the user whose mailbox you want to access.
6. Enter the ID of the message you want to retrieve.
7. The script will retrieve the message properties and display them in the console.
8. If you want to save the message properties to a file, enter "Y" when prompted and follow the instructions.
9. If you want to search for another message, enter "Y" when prompted. Otherwise, enter "N" to exit the script.

## License

This script is licensed under the MIT License. See the LICENSE file for more information.
