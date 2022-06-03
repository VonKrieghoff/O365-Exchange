# O365-Get-MailboxPermission


Two o365 scripts.
One lists all users that have prmissions in someone else mailbox
Other one asks you to input user login name whos permissions you are looking in other user mailboxes.
Script will filter out permissions for users own mailbox, because there is no point for that to show up in results.

Both of them asks user to authenticait against yout o365 cloud, you have to have permissions to read user mailbox permissions to get results by running this script.


# To execute rometely you can run from powershell these single line commands:

## For Single/Specific user permissions:

**$ScriptFromGithHub = Invoke-WebRequest https://raw.githubusercontent.com/VonKrieghoff/O365-Get-MailboxPermission/main/Mailbox%20Permissions%20-%20specific%20user.ps1; Invoke-Expression $($ScriptFromGithHub.Content)**

## For all users that have permissions in someone else mailbox:

**$ScriptFromGithHub = Invoke-WebRequest https://raw.githubusercontent.com/VonKrieghoff/O365-Get-MailboxPermission/main/Mailbox%20Permissions%20-%20all%20users%20except%20self.ps1; Invoke-Expression $($ScriptFromGithHub.Content)**

