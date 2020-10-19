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

.PARAMETER Random
    Generate a random password and set for this user.

.NOTES
    Created by: Josh Crabtree 30 March 20
#>

[CmdletBinding()]

param(
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Enter a name."
        )]
        [ValidateNotNullOrEmpty()]
        [string]$User,

        [Parameter(
            HelpMessage = "Enter a password and click the submit button."
        )]
        [string]$Password,

        #checkbox to make password random
        [Parameter(
            HelpMessage = "Check the box to set a random password for this user."
        )]
        [switch]$MakeRandom
    )    
    begin{
        $FN = "Reset-ADPassword"
        Write-host "Resetting AD Password for $user." -ForegroundColor Cyan

        $Length = 8
        $NonAlphaCharacters = 2
        #.net class for generating a password
        $RandomPassword = [System.Web.Security.Membership]::GeneratePassword($Length, $NonAlphaChars)
        $PassResults = @()
    }
    process{
        $ADUser = Get-ADuser -Identity $User
        if($ADUser){
            if($MakeRandom){
                try{
                    $RandomPassword = -join ((48..57) + (65..90) + (97..122)| Get-Random -Count $Length | foreach {[char]$_})
                    $SetRandomPassword = ConvertTo-SecureString -AsPlainText $RandomPassword -force
                    $ADUser | Set-ADAccountPassword -NewPassword $SetRandomPassword -Reset -ErrorAction Stop
                    Write-Host "Successfully reset the password for $($ADUser) to a random password: $($RandomPassword)" -ForegroundColor Green
                    $Success = $true
                }
                catch{
                    Write-Host "Failed to set the password to a random password | $($error[0].exception.message)." -ForegroundColor Red
                    $Success = $false
                }
                $PassResults += [PSCustomObject] @{
                    Name = $ADUser.name;
                    NewPass = $RandomPassword;
                    Success = $Success;
                }
            }
            else{ 
                try{
                    if($Password){
                        $NewPassword = ConvertTo-SecureString -AsPlainText $Password -Force
                        $ADUser | Set-ADAccountPassword -NewPassword $NewPassword -Reset -ErrorAction Stop
                        Write-host "Successfully reset the password." -ForegroundColor Green
                        $Success = $true
                    }
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
        }
        else{  
            Write-host "User $($User) was not found in AD." -ForegroundColor Red
        }
    }
    end{
        return $PassResults
  }
