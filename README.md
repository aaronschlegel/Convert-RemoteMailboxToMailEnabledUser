# Convert-RemoteMailboxToMailEnabledUser.ps1

This PowerShell script is used to convert a remote mailbox to a mail-enabled user in Office 365. It includes a function called `Remove-O365DirectAssignedLicense` that can be used to remove a direct assigned license from an Office 365 user.

## Usage

To use this script, you need to provide the following parameters:

- `$DomainController`: The domain controller to use for the Active Directory lookup.
- `$AdConnectServer`: The server to use for the Azure AD Connect lookup.
- `$Csvfile`: The CSV file containing the list of remote mailboxes to convert.

You can also use the `Remove-O365DirectAssignedLicense` function to remove a direct assigned license from an Office 365 user. This function takes the following parameters:

- `$userUPNArray`: An array of user UPNs to remove the license from.
- `$userUPNFilePath`: The path to a file containing a list of user UPNs to remove the license from.
- `$user`: The UPN of the user to remove the license from.
- `$processAllUsers`: A switch parameter that indicates whether to remove the license from all users.

## Requirements

This script requires PowerShell version 5.1 or later and must be run as an administrator. It also requires the Exchange Management Shell to be installed on the server where it will be run.

## CSV File Format

The CSV file should have two columns: `SourceUPN` and `TargetMailbox`. The `SourceUPN` column should contain the UPN of the remote mailbox to convert, and the `TargetMailbox` column should contain the forwarding address for the mail-enabled user.

