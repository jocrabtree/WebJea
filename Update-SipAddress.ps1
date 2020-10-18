<#  
.SYNOPSIS
    This function updates a user's sip address in active directory

.DESCRIPTON
    Provide a username and the 'msrtcsip-primaryaddress' ad attribute is update to the primary smtp address.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> Update-SipAddress -user <name>

.PARAMETER User
    Enter the samaccount as a string.

.NOTES
    Created by: Josh Crabtree (@uc_crab) 14 April 20  
#>

[CmdletBinding()]
param(
        [Parameter(
            Mandatory=$True,
            Position=0,
            ValueFromPipeline=$True
        )]
        $User
    )
    begin{
        "Updating SIP based on Primary SMTP for $User" | Write-Host -ForegroundColor Cyan
        $Results = @()
    }
    process{
        $ADUser = Get-ADuser -Identity $User -properties 'msrtcsip-primaryuseraddress',mail
        $SIP = $ADUser."msrtcsip-primaryuseraddress"
        $Mail = $ADUser.mail
        if($SIP -and "sip:$Mail" -ne $SIP){ #if both the sip address and primary email DO NOT match, try to update
            try{
                $ADUser | Set-ADUser -Replace @{"msrtcsip-primaryuseraddress"="sip:$Mail"} -ErrorAction Stop #actual update of sip address in AD happens on this line
                "Successfully Updated SIP Address." | Write-Host -ForegroundColor Green
                $Updated = $True
            }
            catch{#catch the error
                "Failed to Update SIP Address| $($error[0].exception.message)." | Write-Host -ForegroundColor Red
                $Updated = $False
            }
            $Results += [pscustomobject] @{
                Name=$ADUser.Name;
                Mail=$Mail;
                OldSIP=$SIP;
                NewSIP="sip:$Mail";
                Updated=$Updated;
            }
        }
        else{ #else for the if statement above. if the mail and sip match, skip this user.
            "SIP already matches Primary SMTP. No updates made." | Write-Host -ForegroundColor Cyan
        }
    }
    end{
        return $Results
   }
