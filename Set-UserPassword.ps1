<#  
.SYNOPSIS
    This function resets a user's password in AD.

.DESCRIPTON
    Provide a username and password to the script and it will reset a user's password.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> Set-UserPassword -user <name> -password <password>

.PARAMETER User
    Enter the samaccount as a string.

.PARAMETER Password
    Enter the password for this user.

.NOTES
    Created by: Josh Crabtree 30 March 20 
#>

[CmdletBinding()]

param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Enter a name."
            #ValueFromPipeline = $true
        )]
        [ValidateNotNullOrEmpty()]
        [string]$User,

        [Parameter(
            Mandatory= $true,
            HelpMessage = "Enter a password and click the submit button."
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Password
    )    
    begin{
        $FN = "Reset-ADPassword"
        Write-host "Resetting AD Password for $user." -ForegroundColor Cyan
        $NewPassword = ConvertTo-SecureString -AsPlainText $Password -Force
        $PassResults = @()
    }
    process{
        $ADUser = Get-ADuser -Identity $User
        if($ADUser){
            try{
                #command in AD to set password
                $ADUser | Set-ADAccountPassword -NewPassword $NewPassword -Reset -ErrorAction Stop
                Write-host "Successfully reset the password." -ForegroundColor Green
                $Success = $true
            }
            catch{
                Write-host "Failed to reset the password | $($error[0].exception.message)." -ForegroundColor Red
                $Success = $false
            }
            $PassResults += [PSCustomObject] @{
                    Name=$ADUser.name;
                    NewPass=$Password;
                    Success=$Success;
            }
        }
        #else statement for if above. if we can't find a user, return not found.
        else{
            Write-host "User $($User) was not found in AD." -ForegroundColor Red
        }
    }
    end{
        return $PassResults
  }
