#Remote invoke:
#$ScriptFromGithHub = Invoke-WebRequest https://raw.githubusercontent.com/VonKrieghoff/O365-Get-MailboxPermission/main/Mailbox%20Permissions%20-%20all%20users%20except%20self.ps1; Invoke-Expression $($ScriptFromGithHub.Content)


### Description
# 1. Script connects to your o365 tennant
# 2. looks for mailboxes
# 3. Outputs mailboxes where someone else have some permissions except user it self.

Clear-Host

        Write-Host "








              __     __
             /  \~~~/  \
       ,----(     ..    )
      /      \__     __/
     /|         (\  |(       
    ^ \   /___\  /\_|        
       |__|   |__|           
    
       " -ForegroundColor Yellow
    Write-Host "I'm Reading Mailbox Permissions
    
    
    " -ForegroundColor Yellow


#Module needed for o365 connection
Import-Module ExchangeOnlineManagement

#Creating connection to o365, another windows will popup asked for login credentials to access your o365 tenant. Show banner hides built in info output at connection.
Write-Progress -Activity "Waiting for you to log in o365"
Connect-ExchangeOnline -ShowBanner:$false
# You can also use "Connect-ExchangeOnline -UserPrincipalName myloginname@mydomain.com" to prefill form so only password and 2FA will be asked.


#EXPORT TO Excel
#EXPORT TO Excel
#EXPORT TO Excel
Install-Module ImportExcel -Scope CurrentUser -ErrorAction SilentlyContinue #Will install import excel module, otherwise it will not be possible to export to excel.
$date = (get-date -UFormat "%Y-%m-%d (%H-%M-%S)") #Gets date and time for excel file name.



$ExcelFileName = "O365-Get-MailboxPermission" #Excel and folder name
########### FOLDER
$FolderName = "c:\TEMP\$ExcelFileName"
if (Test-Path $FolderName) {
    Write-Host "Folder Exists: " -NoNewline
    Write-Host "$FolderName
    
    "  -ForegroundColor Yellow
    # Perform Delete file from folder operation
}
else
{
    #PowerShell Create directory if not exists
    Write-Host "Folder $FolderName, was not found, Creating ......."  -ForegroundColor Yellow
    New-Item $FolderName -ItemType Directory  | Out-Null
    Write-Host "Folder Created successfully - "
    Write-Host "$FolderName
    
    "  -ForegroundColor Yellow

}
########### FOLDER

#$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition #Detects current folder frome where are you executing script, if localy execute
$scriptPath = "$FolderName"

$ExcelFile = "$scriptPath\$ExcelFileName-$date.xlsx" #Genereates excel file location and name

#EXPORT TO Excel
#EXPORT TO Excel
#EXPORT TO Excel



Write-Host "Output result file will be saved in: " -nonewline
Write-Host "$ExcelFile 

"-ForegroundColor Yellow


Write-Progress -Activity "For ~5000 users it takes about 10min"

$DateStart = Get-Date
Write-Host "$DateStart " -nonewline
Write-Host "Running ........."-ForegroundColor Yellow
## To output in console:
#Get-EXOMailbox -ResultSize 1000 | select-object @{n='Identity';e={$_.UserPrincipalName}} | Get-MailboxPermission | Where-Object { -not ($_.User -like "NT AUTHORITY\SELF") } | format-table -AutoSize

## To Output in Excel:
Get-EXOMailbox -ResultSize 100 | select-object @{n='Identity';e={$_.UserPrincipalName}} | Get-MailboxPermission | Where-Object { -not ($_.User -like "NT AUTHORITY\SELF") } | Export-Excel $ExcelFile -AutoSize -StartRow 2 -TableName Report
# Get-Mailbox -ResultSize unlimited - gets all mailboxes in o365 tenant, you can also replace unlimited with 1000 for example so only 1000 mailboxes will be red.
# select-object @{n='Identity';e={$_.UserPrincipalName}} - Maps UserPrincipalName as Identity, this is needed because if you have duplicated user Full names in directory the error will happen and results will not look clean.
# Get-MailboxPermission | Where-Object { -not ($_.User -like "NT AUTHORITY\SELF") } - gets mailbox permissions except where user have permissions for its own mailbox, there is no point of that information, of course user will have access to its own mailbox.
# format-table -AutoSize - formats output table with dynamic column width

$DateEND = Get-Date
Write-Host "$DateEND " -nonewline
$DURATION = $DateEND - $DateStart
$DURATIONmin = $DURATION.Minutes
$DURATIONSec = $DURATION.Seconds

Write-Host "DONE 
" -ForegroundColor Green
Write-Host "Finished in " -ForegroundColor Green -nonewline
Write-Host "$DURATIONmin" -ForegroundColor RED -nonewline
Write-Host " minutes" -ForegroundColor Green -nonewline
Write-Host " and " -ForegroundColor Green -nonewline
Write-Host "$DURATIONSec" -ForegroundColor RED -nonewline
Write-Host " seconds.

" -ForegroundColor Green

Write-Host "In Output file IDENTITY column is target mailbox, where user from USER column have permissions to access it

"  -ForegroundColor Yellow



Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue

