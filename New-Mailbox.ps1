<#  
.SYNOPSIS
    This script creates a new remote mailbox for a user in a hybrid Exchange environment.

.DESCRIPTON
    Provide a username (samaccount) and click the submit button to create a new mailbox for a user account.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> .\New-Mailbox.ps1 -users <usernames>

.PARAMETER Users
    Enter each samaccount as a string.

.NOTES
    Created by: Josh Crabtree 2 Sept 2020
#>

[CmdletBinding()]
 
param(
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Enter one username per line (Spelling is crucial!). Samaccount is used for the username."
    )]
    [ValidateNotNullOrEmpty()]
    #create mailboxes for up to 10 users at once
    [ValidateCount(1,10)] 
    [string[]]$users
)
begin{
    $FN = "New-Mailbox"
    Write-Host "$FN | Creating Mailboxes for New Users."
    #replace <YOUR_DOMAIN_HERE> with you domain. ex: "contoso.mail.onmicrosoft.com"
    $RemoteSuffix = "@<YOUR_DOMAIN_HERE>.mail.onmicrosoft.com"
    $Results=@()
    $MailboxOnPrem=@()
    $Today = (Get-Date).ToString('MM-dd-yyyy')
    $ExactTime = Get-Date -Format "MM-dd-yyyy HHmm tt"
     #feel free to change to a log file location you may already have
    $logfile = "C:\Logs\NewMailbox-$($Today).log"
    #use suffix for the UPN
    $suffix = '@<YOUR_DOMAIN>.com'
     #feel free to change to a location where a csv of all mailboxes created are kept
    $TotalMBCreated = "\\<YOUR_LOCATION>\NewMailboxes.csv"
        
    #User account and password for creds
    #update this line with your account creds. ex: "admin@contoso.com"
    $UN = "<YOUR_ACCOUNT_NAME>@<YOUR_DOMAIN>.com" 
    #use a AES key
    $Key = Get-Content "C:\Admin\AES.key" 
    #convert the password for your admin account to a securestring with the AES key
    $PW = Get-Content "C:\Admin\Password.txt" | ConvertTo-SecureString -Key $Key 
    #build the creds with the username and password
    $Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UN, $PW 

    #Connect Exchange OnPrem
    $EXFN = "Connect-ExchangeOnPrem"
    try{
        #update with your on-prem mail server name & domain. ex: mailserver1.contoso.com
        $EOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<YOUR_MAIL_SERVER>.<YOUR_DOMAIN>/PowerShell/ -Authentication Kerberos -Credential $Creds -AllowRedirection 
        $importex = Import-PSSession $EOPSession -AllowClobber -WarningVariable ignore -InformationAction Ignore -DisableNameChecking

        $Message = "$EXFN | Successfully connected to Exchange OnPrem"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
        
    }
    catch{
        $Message="$EXFN | Failed to connect to Exchange OnPrem! | $($error[0].exception.message)"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }

}
process{
    foreach($User in $Users){
        $ADUser = Get-ADUser -identity $User -properties GivenName,Surname,Mail
        $Prefix = $ADUser.samaccountname
        $UPN = $ADUser.UserPrincipalName
        if($UPN -like "*$Suffix"){
            # If mailbox exists on-prem only, it will be logged and script will not attempt to create in cloud.
            if(Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue){
                $Message = "$FN | $($ADUser.name) has an onprem mailbox already, please fix and try again"
                Write-Host $Message
                $MailboxOnPrem += $ADUser
            }
            else{
                $Message = "$FN | Adding Mailbox for $UPN"
                Write-Host $Message
                try{
                    #'enable-remotemailbox' is run in on-prem exchange to create a mailbox in the cloud with a reference on-prem.
                    Enable-RemoteMailbox -Identity $UPN -RemoteRoutingAddress (-join ($Prefix,$RemoteSuffix)) -Confirm:$False 
                    $Message = "$FN | Mailbox created for user $($ADUser.name) in on-prem and O365."
                    Write-Host $Message
                    #wait 5 seconds before checking mailbox exists
                    sleep 5
                    $NewMB = Get-RemoteMailbox $ADUser.UserPrincipalName
                    $MailboxMade = "TRUE"
                    $Mail = $NewMB.PrimarySMTPAddress
                    "$($ADUser.Name) - $($NewMB.PrimarySMTPAddress) - $(Get-Date)" | Out-File $TotalMBCreated -Append
                }
                catch{
                    $Message = "$FN | Failed to create a remote mailbox for user $($ADUser.name) in on-prem and O365. | $($error[0].exception.message)."
                    Write-Host $Message 
                    $MailboxMade = "FALSE"
                    $Mail = "N/A"
                }
            }
        }
        #else statement for if above. if the user doesn't have a upn, then skip this user.
        else{ 
            $Message = "There was an issue with the UPN $UPN and the mailbox was not created."
            Write-Host $Message
            $MailboxMade = "FALSE"
            $Mail = "N/A"
        }
        $Results += [pscustomobject] @{
            Name=$ADUser.Name;
            Mail=$Mail;
            MailboxMade=$MailboxMade;
        }
    }
}
end{
    Write-host ""
    Write-Host "$FN | Mailbox Creation Results:"
    Write-Host ""
    $Results
    if($MailboxOnPrem){
        Write-Host ""
        Write-Host "$FN | Users with Mailbox On-Prem:"
        Write-Host ""
        $MailboxOnPrem.ForEach{"$($MailboxOnPrem.IndexOf($_) + 1)) | $($_.Name) | $($_.mail)"}
    }
}