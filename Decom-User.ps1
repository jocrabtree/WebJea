<#  
.SYNOPSIS
    This script decoms a user based on options selected via the checkboxes below.

.DESCRIPTON
    Provide a username (samaccount) and select the options to decom the account.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> .\Decom-User.ps1 -names <names>

.PARAMETER names
    Enter each samaccount as a string.

.NOTES
    Created by: Josh Crabtree & Mike Bratton 13 May 20 
#>

[CmdletBinding()]

param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Enter one username per line (Spelling is crucial!). Samaccount is used for the username."
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateCount(1,10)]
        [string[]]$Names,

        [Parameter(
            HelpMessage = "Check the box to use all options to decom the user."
        )]
        [switch]$SelectAll,

        #have not added code for this yet, placeholder for future improvement
        [Parameter(
            HelpMessage = "Check the box to set an OOO Message for both Internal & External Mail Recipients."
        )]
        [switch]$OOOInternalExternal,

        #have not added code for this yet, placeholder for future improvement
        [Parameter(
            HelpMessage = "Check the box to set an OOO Message for Internal Mail Recipients only."
        )]
        [switch]$OOOInternalOnly,

        #have not added code for this yet, placeholder for future improvement
        [Parameter(
            HelpMessage = "Check the box to set an OOO Message for External Mail Recipients only."
        )]
        [switch]$OOOExternalOnly,

        [Parameter(
            HelpMessage = "Check the box to disable the user's Skype for Business account."
        )]
        [switch]$DisableSkypeAccount,

        [Parameter(
            HelpMessage = "Check the box to remove O365 licensing."
        )]
        [switch]$RemoveO365Licensing,

        [Parameter(
            HelpMessage = "Check the box to remove all scheduled meetings."
        )]
        [switch]$RemoveMeetings,

        [Parameter(
            HelpMessage = "Check the box to remove the user from all AD groups."
        )]
        [switch]$RemoveADGroups,

        [Parameter(
            HelpMessage = "Check the box to remove the user from all O365 groups."
        )]
        [switch]$Remove365Groups
    )  
    
begin{
    $FN = "Process-Termed"
    $Today = (Get-Date).ToString('MM-dd-yyyy')
    $ExactTime = Get-Date -Format "MM-dd-yyyy HHmm tt"
    $logfile = "C:\Logs\ProcessedTermed-$($Today).log"
    $suffix = '*<YOUR_DOMAIN>.com' #ex: '*contoso.com'
    

    #User account and password for creds
    
    #ex: "admin@contoso.com"
    $UN = "<YOUR_ACCOUNT>@<YOUR_DOMAIN>.com" 
    #AES Key
    $Key = Get-Content "C:\Admin\AES.key"
    #convert password to secure string
    $PW = Get-Content "C:\Admin\Password.txt" | ConvertTo-SecureString -Key $Key
    #build creds from username and password above
    $Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UN, $PW
    
    #Define global variable used by the script
    $DomainGroup = (Get-ADGroup -f{name -like "Domain Users"} -Properties primaryGroupToken)
    #I use a 'terminated users' group in AD
    $TermedGroup = (Get-ADGroup -f{name -like "Terminated*"} -Properties primaryGroupToken)
    $TermedGroupID = $TermedGroup.PrimaryGroupToken
    #ex: "contoso:EMS"
    $EMS = "<YOUR_DOMAIN>:EMS"
    #ex: "contoso:WIN_DEF_ATP"
    $WinDef = "<YOUR_DOMAIN>:WIN_DEF_ATP"
    #ex: "contoso:ENTERPRISEPACK"
    $E3 = "<YOUR_DOMAIN>:ENTERPRISEPACK"
    #ex: "contoso:ENTERPRISEPREMIUM"
    $E5 = "<YOUR_DOMAIN>:ENTERPRISEPREMIUM"
    $SelectAllResults = @()
    $SkypeResults = @()
    $365GroupResults = @()
    $ADGroupResults = @()
    $licensingResults = @()
    $MeetingResults = @()
    $ToProcess = @()
    $Properties = "telephoneNumber","Displayname","UserPrincipalName","title","department","mail","msExchMailboxGuid","manager","physicalDeliveryOfficeName","msRTCSIP-UserEnabled","msRTCSIP-PrimaryUserAddress", "msExchWhenMailboxCreated", "distinguishedName", "extensionAttribute11","sAMAccountName","extensionAttribute5","extensionAttribute8","enabled","MemberOf","msExchRecipientTypeDetails","msRTCSIP-DeploymentLocator"
   
    $TermedMonth = "$($(Get-Date).ToString("MMM").ToUpper())"
    $TermedYear = "$($(Get-Date).ToString("yyyy"))"
    #make sure to update with <YOUR_OU> & <YOUR_DOMAIN>
    $TermedOU = "OU=$($TermedMonth),OU=$($TermedYear),OU=Terminated,OU=Users,OU=<YOUR_OU>,DC=<YOUR_DOMAIN>,DC=com"


    #Connect Exchange Online
    $EXFN = "Connect-ExchangeOnline"
    try{
        $EOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection 
        $Import = Import-PSSession $EOLSession -AllowClobber -WarningVariable ignore -InformationAction Ignore -DisableNameChecking
        
        $Message = "$EXFN | Successfully connected to Exchange Online"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
        
    }
    catch{
       $Message="$EXFN | Failed to connect to Exchange Online! | $($error[0].exception.message)"
       Write-Host $Message
       "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }

    #Connect to O365
    $365FN = "Connect-Office365"
    try{
        Connect-MsolService -Credential $Creds
        $Message = "$365FN | Connected to Microsoft Office 365"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }
    catch{
        $Message = "$365FN | Failed to Microsoft Office 365! | $($error[0].exception.message)"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }

    #Connect to S4B On-Prem
    $SFBFN = "Connect-SFBOnPrem"
    try{
        $sessionOpt = New-PSSessionOption  -SkipRevocationCheck
        #change <YOUR_S4B_SERVER> & <YOUR_DOMAIN>
        $session = New-PSSession -ConnectionUri https://<YOUR_S4B_SERVER>.<YOUR_DOMAIN>.com/ocsPowerShell/ -Credential $creds -SessionOption $sessionOpt
        $importsfb = Import-PSSession $session -AllowClobber -WarningVariable ignore -InformationAction Ignore -ErrorAction Stop
        $Message = "$SFBFN | Connected to Skype for Business On-Prem"
        Write-Host $Message
         "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }
    catch{
        $Message = "$SFBFN | Failed to connect to Skype for Business On-Prem! | $($error[0].exception.message)"
        Write-Host $Message
        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
    }
    Write-Host ""
    Write-Host "$FN | Processing termination for requested users." | Out-File -FilePath $logfile
    $Names.ForEach{Write-Host "$($Names.IndexOf($_) + 1)) $($_)"}
}
process{
    foreach($name in $names){
        write-host ""
        $user = $null
        try{
            $User = Get-ADUser -identity $name -Properties * -ErrorAction stop
        }
        catch{
            $Message = "$FN | $Name not found in Active Directory. Skipping this user."
            Write-Host $Message
            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append  
        }
        $UPN = $User.UserPrincipalName
        if($User){
            $Message = "$FN | Processing User $($User.Name)"
            Write-Host $Message
            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File -filepath $logfile -Append
            $ToProcess += $User
            if($SelectAll){
                
                # Activating All Switches
                $OOOInternalExternal=$True;$OOOInternalOnly=$True;$OOOExternalOnly=$True;$DisableSkypeAccount=$True;$RemoveExchMailbox=$True;$RemoveO365Licensing=$True;$RemoveMeetings=$True;$RemoveADGroups=$True;$Remove365Groups=$True

                # Set Account to Terminated and Disable
                $User | Set-ADUser -Enabled $False -Replace @{ExtensionAttribute8="TERMINATED"}
                
                # Add to Terminated Users Group and Set as Primary
                $Message = "$FN | Adding to Terminated Users Group"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                try{
                    Add-ADGroupMember -Identity $TermedGroup -Members $User -Confirm:$False -ErrorAction Stop
                    $Message = "$FN | Successfully added $($User.name) to the $($TermedGroup.Name) group."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                catch{
                    $Message = "$FN | Failed to add $($User.name) to the $($TermedGroup.Name) group. | $($error[0].exception.message)"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                $Message = "$FN | Adding Terminated Users Group as Primary"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                try{
                    $User | Set-ADUser -Replace @{primarygroupID=$TermedGroupID} -ErrorAction Stop
                    $Message = "$FN | Successfully set the $($TermedGroup.Name) group as the primary group."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                catch{
                    $Message = "$FN | Failed to set the $($TermedGroup.Name) group as the primary group. | $($error[0].exception.message)"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                # Remove Domain Users Group
                $Message = "$FN | Removing Domain Users Group"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                try{
                    Remove-ADGroupMember -Identity $DomainGroup -Members $User -Confirm:$False
                    $Message = "$FN | Successfully removed $($User.name) from the $($DomainGroup.Name) group."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                catch{
                    $Message = "$FN | Failed to remove $($User.name) from the $($DomainGroup.Name) group. | $($error[0].exception.message)"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
            }

            # Remove from all AD Groups
            if($RemoveADGroups){
                $Groups = [pscustomobject]@{
                    Removed=@();
                    NotRemoved=@();
                }
                $Message = "$FN | Processing AD Group Removal"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                $User.MemberOf | ForEach-Object {
                    $Group = Get-ADGroup $_
                    $Message = "$FN | Removing $($User.Name) from the group $($Group.Name)"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    try{
                        $Group | Remove-ADGroupMember -Members $User.DistinguishedName -Confirm:$false -ErrorAction Stop
                        $Message = "$FN | Successfully removed $($User.name) from the group $($Group.Name)."
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Groups.Removed += $Group.Name
                    }
                    catch{
                        $Message = "$FN | Failed to remove $($User.name) from the group $($Group.Name). | $($error[0].exception.message)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Groups.NotRemoved += $Group.Name
                    }
                }
                $ADGroupResults += [pscustomobject] @{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    RemovedGroups=($Groups.Removed);
                    NotRemovedGroups=($Groups.NotRemoved);
                }
            }

            # Check for O365 Account
            if($SelectAll -or $OOOInternalExternal -or $OOOInternalOnly -or $OOOExternalOnly -or $RemoveO365Licensing -or $RemoveMeetings -or $Remove365Groups){
                $Message = "$FN | Checking for an O365 account"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                try{
                    $MSOLUser = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
                    $Message = "$FN | $($MSOLUser.DisplayName) has an O365 account."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
                catch{
                    $Message = "$FN | $($User.name) does not have an O365 account."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                }
            }

            # Remove From All O365 Groups in Cloud
            if($Remove365Groups){
                $OGroups = [pscustomobject]@{
                    Removed=@();
                    NotRemoved=@();
                    RemovedOwner=@();
                    NotRemovedOwner=@();
                }
                if($MSOLUser){
                    $O365Groups = @()
                    $User.memberof.foreach{$GroupName = $_.split(",")[0].replace("CN=","");if($GroupName -cmatch "Group_"){$O365Groups += $GroupName.replace("Group_","")}}
                    $UserEmail = ($MsolUser.ProxyAddresses | where {$_ -cmatch "SMTP" }).replace("SMTP:","")
                    $Message = "$FN | Removing from O365 Groups"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    if($O365Groups){
                        $O365Groups | Foreach{
                            $MsolGroup = Get-MsolGroup -ObjectId $_
                            $UnifiedGroup = Get-UnifiedGroup -Identity $MsolGroup.DisplayName
                            if($UnifiedGroup){
                                $Message = "$FN | Processing the $($UnifiedGroup.DisplayName) group..."
                                Write-Host $Message
                                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                $Owners = Get-UnifiedGroupLinks -Identity $UnifiedGroup.DisplayName -LinkType Owners
                                if($Owners.PrimarySmtpAddress -contains $UserEmail){
                                    $Message = "$FN | $($MSOLUser.DisplayName) is an owner of this group, removing owner status..."
                                    Write-Host $Message
                                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                    try{
                                        Remove-UnifiedGroupLinks $UnifiedGroup.DisplayName -LinkType Owners -Links $UserEmail -Confirm:$false
                                        $Message = "$FN | Successfully removed $($MSOLUser.DisplayName) as owner of the $($UnifiedGroup.DisplayName) O365 group."
                                        Write-Host $Message
                                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                        $OGroups.RemovedOwner += $UnifiedGroup.DisplayName

                                    }
                                    catch{
                                        $Message = "$FN | Failed to remove $($MSOLUser.DisplayName) as owner of the $($UnifiedGroup.DisplayName) O365 group. | $($error[0].exception.message)"
                                        Write-Host $Message
                                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                        $OGroups.NotRemovedOwner += $UnifiedGroup.DisplayName
                                    }
                                    sleep 30
                                }
                                try{
                                    Remove-UnifiedGroupLinks $UnifiedGroup.DisplayName -LinkType Members -Links $UserEmail -Confirm:$false -ErrorAction Stop
                                    $Message = "$FN | Successfully removed $($MSOLUser.DisplayName) from the $($UnifiedGroup.DisplayName) O365 group."
                                    Write-Host $Message
                                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                    $OGroups.Removed += $UnifiedGroup.DisplayName
                                }
                                catch{
                                    $Message = "$FN | Failed to remove $($MSOLUser.DisplayName) from the $($UnifiedGroup.DisplayName) O365 group. | $($error[0].exception.message)"
                                    Write-Host $Message
                                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                    $OGroups.NotRemoved += $UnifiedGroup.DisplayName
                                }
                            }
                        }
                    }
                    else{
                        $Message = "$FN | There are no 365 Groups to remove from $($MSOLUser.DisplayName)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    }
                }
                $365GroupResults += [pscustomobject] @{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    RemovedOGroups=($OGroups.Removed);
                    NotRemovedOGroups=($OGroups.NotRemoved);
                    RemovedOGroupsOwner=($OGroups.RemovedOwner);
                    NotRemovedOGroupsOwner=($OGroups.NotRemovedOwner);
                }
            }

            # Remove all O365 Licenses from User
            If($RemoveO365Licensing){
                $Licenses = [pscustomobject]@{
                    Removed=@();
                    NotRemoved=@();
                }
                if($MSOLUser){
                    $Message = "$FN | Removing O365 Licensing"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    $MSOLUserLicenses = $MSOLUser.licenses.AccountSkuID
                    if($MSOLUserLicenses){
                        $MSOLUserLicenses | ForEach-Object {
                            if(($_ -ne $EMS) -and ($_ -ne $E5) -and ($_ -ne $E3) -and ($_ -ne $WinDef)){
                                $LicenseName = $_.split(":")[1]
                                $Message = "$FN | Removing the $LicenseName license from $($MSOLUser.DisplayName)."
                                Write-Host $Message
                                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                try{
                                    Set-MsolUserLicense -UserPrincipalName $MSOLUser.UserPrincipalName -RemoveLicenses $_ -ErrorAction Stop
                                    $Message = "$FN | Successfully removed the $LicenseName license from $($MSOLUser.DisplayName)."
                                    Write-Host $Message
                                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                    $Licenses.Removed += $LicenseName
                                }
                                catch{
                                    $Message = "$FN | Failed to remove the $LicenseName license from $($MSOLUser.DisplayName). | $($error[0].exception.message)"
                                    Write-Host $Message
                                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                                    $Licenses.NotRemoved += $LicenseName
                                }
                            }
                            else{
                                $Message = "$FN | $($_.Split(":")[1]) was removed with its licensing group."
                                Write-Host $Message
                                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            }
                        }
                    }
                    else{
                        $Message = "$FN | There are no licenses to remove from $($MSOLUser.DisplayName)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    }
                }
                $LicensingResults += [pscustomobject] @{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    RemovedLicenses=($Licenses.Removed)
                    NotRemovedLicenses=($Licenses.NotRemoved);
                }
            }

            # Remove User Mailbox Meetings
            if($RemoveMeetings){
                $Message = "$FN | Checking for Meetings in Exchange"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                if($UPN -like $Suffix){
                    $Mailbox = Get-Mailbox $UPN
                    $Meetings = Remove-CalendarEvents -Identity $Mailbox.DistinguishedName -CancelOrganizedMeetings -QueryWindowInDays 160 -Confirm:$False -PreviewOnly
                    if($Meetings){
                        $Message = "$FN | $($User.Name) Removing all scheduled meetings with attendees"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Meetings.ForEach{
                            $Message = "$FN | $_"
                            Write-Host $Message
                            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        }
                        try{
                            Remove-CalendarEvents -Identity $Mailbox.DistinguishedName -CancelOrganizedMeetings -QueryWindowInDays 160 -Confirm:$False
                            $Message = "$FN | $($User.Name) | Removed all meetings"
                            Write-Host $Message
                            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            $RemovedMeetings = $True
                        }
                        catch{
                            $Message = "$FN | $($User.Name) | Failed to remove all meetings."
                            Write-Host $Message
                            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            $RemovedMeetings = $False
                        }
                        $MeetingNames = $($Meetings.foreach{$_.split('"')[1]}) -join ","
                    }
                    else{
                        $Message = "$FN | $($User.Name) | There are no meetings to remove"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $RemovedMeetings = "N/A"
                    }
                }
                else{
                    $Message = "$FN | $($User.Name) | Does not have a valid UserPrincipalName"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    $RemovedMeetings = "No UPN, Aborted!!"
                }
                $MeetingResults += [pscustomobject] @{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    Meetings=$MeetingNames;
                    MeetingsRemoved=$RemovedMeetings;
                }
            }

            # Disable SFB On-Prem
            if($DisableSkypeAccount){
                if($UPN -like $Suffix){
                    $CSUser = Get-CsUser -Identity $UPN -ErrorAction SilentlyContinue
                    if($CSUser){
                        $Message =  "$FN | Disabling Skype for Business access"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        try{
                            $CSUser | Disable-CsUser -ErrorAction Stop
                            $Message = "$FN | Successfully disabled SFB On-Prem for $($User.Name)."
                            Write-Host $Message
                            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            $Disabled = "True"
                            try{
                                $User | Set-ADUser -Clear "msRTCSIP-UserEnabled","msRTCSIP-DeploymentLocator","msRTCSIP-FederationEnabled","msRTCSIP-InternetAccessEnabled","msRTCSIP-Line","msRTCSIP-OptionFlags","msRTCSIP-PrimaryHomeServer","msRTCSIP-PrimaryUserAddress","msRTCSIP-UserPolicies","msRTCSIP-UserRoutingGroupID"
                                $Message = "$FN | Successfully removed Enablement and SIP address for $($User.Name)."
                                Write-Host $Message
                                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            }
                            catch{
                                $Message = "$FN | Failed to remove Enablement and SIP address for $($User.Name). | $($error[0].exception.message)"
                                Write-Host $Message
                                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            }
                        }
                        catch{
                            $Message = "$FN | Failed to disable SFB On-Prem for $($User.Name). | $($error[0].exception.message)"
                            Write-Host $Message
                            "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                            $Disabled = "False"
                        }
                    }
                    else{
                        $Disabled = "Already Disabled"
                    }
                }
                else{
                    $Message = "$FN | $($User.Name) does not have a UserPrincipalName! CSUser was not attempted to be found."
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    $Disabled = "No UPN, Aborted!!"
                }
                $SkypeResults += [pscustomobject] @{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    SFBDisabled=$Disabled;
                }
            }

            # Complete Termination of User if Select All
            if($SelectAll){
                # Remove Manager 
                if($User.Manager){
                    $Message = "$FN | Removing Manager"
                    Write-Host $Message
                    try{
                        $Manager = $User.Manager.split(",")[0].replace("CN=","")
                        $User | Set-ADUser -Clear "Manager"
                        $Message = "$FN | Successfully removed the manager $Manager from $($User.name)."
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Removed = "True"
                    }
                    catch{
                        $Message = "$FN | Failed to remove the manager $Manager from $($User.name). | $($error[0].exception.message)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Removed = "False"
                    }
                }
                else{
                    $Message = "$FN | $($User.name) does not have a Manager set"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    $Manager = "No Manager Set"
                    $Removed = "N/A"
                }

                # Remove Phone Number
                if($User.telephoneNumber){
                    $TelephoneNumber = $User.telephoneNumber
                    $Message = "$FN | Removing Telephone Number"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    try{
                        $User | Set-ADUser -Clear "telephoneNumber"
                        $Message = "$FN | Successfully removed the telephone number $($User.telephonenumber) for $($User.name)."
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Phone = "True"
                    }
                    catch{
                        $Message = "$FN | Failed to remove the telephone number $($User.telephonenumber) for $($User.name). | $($error[0].exception.message)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Phone = "False"
                    }
                }
                else{
                    $Message = "$FN | $($User.name) does not have a telephone number"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    $Phone = "No Phone"
                }

                # Set Description to Offboard Date
                $Message = "$FN | Setting Description to Offboarded Date"
                Write-Host $Message
                "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                $User | Set-ADUser -Description "Deprovisioned $ExactTime via WebJea Decom."

                # Move User to Termed OU
                if($User.distinguishedname -notlike $TermedOUs){
                    $Message = "$FN | Moving User to the Termed OU"
                    Write-Host $Message
                    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                    try{
                        Move-ADObject -Identity $user.DistinguishedName -TargetPath $TermedOU -ErrorAction Stop
                        $Message = "$FN | Successfully moved $($User.name) to the Termed OU ($TermedOU)."
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Moved = "True"
                    }
                    catch{
                        $Message = "$FN | Failed to move $($User.name) to the Termed OU. | $($error[0].exception.message)"
                        Write-Host $Message
                        "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append
                        $Moved = "False"
                    }
                }
                $SelectAllResults += [pscustomobject]@{
                    Name=$User.Name;
                    UserPrincipalName=$User.UserPrincipalName;
                    Title=$User.Title;
                    Department=$User.Department;
                    DistinguishedName=$User.DistinguishedName;
                    Moved=$Moved;
                    Manager=$Manager;
                    ManagerRemoved=$Removed;
                    TelephoneRemoved=$Phone;
                    TelephoneNumber=$TelephoneNumber
                    ReasonForRemoval=("Terminated and Disabled");
                    RemovedGroups=($Groups.Removed);
                    NotRemovedGroups=($Groups.NotRemoved);
                    RemovedOGroups=($OGroups.Removed);
                    NotRemovedOGroups=($OGroups.NotRemoved);
                    RemovedOGroupsOwner=($OGroups.RemovedOwner);
                    NotRemovedOGroupsOwner=($OGroups.NotRemovedOwner);
                    Meetings=$MeetingNames;
                    MeetingsRemoved=$RemovedMeetings;
                    RemovedLicenses=($Licenses.Removed)
                    NotRemovedLicenses=($Licenses.NotRemoved);
                    SFBDisabled=$Disabled;
                }
            }
        }
    }
}
end{
    Write-host ""
    Write-Host "$FN | Deprovisioning complete for the following users"
    $ToProcess.ForEach{"$($ToProcess.IndexOf($_) + 1)) $($_.Name)"}
    If($SelectAllResults){
        $SelectAllResults | Export-csv C:\Logs\Decom\DecomUsers-$Today.csv -Append
    }
    #kill any open sessions
    Get-PsSession | Remove-PSSession
}
