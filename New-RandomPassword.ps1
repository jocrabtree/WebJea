<#  
.SYNOPSIS
    Click the submit button to generate a random eight character password.

.DESCRIPTON
    Click the submit button to generate a random eight character password.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> New-RandomPassword -Length <int> -ExcludeSpecialCharacters

.PARAMETER Length
    NOTE: Enter the length as a positive number.

.PARAMETER ExcludeSpecialCharacters
    NOTE: Use this checkbox to exclude any special characters
    (Ex: Special characters: ! " # $ % & * + , < == => =? @)

.NOTES
    Created by: Josh Crabtree 24 April 2020 
#>

[CmdletBinding()]
    
param(
        [Parameter(
            Mandatory= $true
            #HelpMessage = "Enter the number of characters you want in the password."
        )]
        [ValidateNotNullorEmpty()]
        [string]$Length,
        
        #checkbox in WebJea is created via a switch statement. this checkbox will exclude special characters.
        [Parameter()]
        [switch]$ExcludeSpecialCharacters
    )
 
 
    BEGIN {
        $nonAlphaCharacters = 2
        #.net class for generate password
        $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
        #special characters to exclude
        $SpecialCharacters = @((33,35) + (36..38) + (42..44) + (60..64))
    }
 
    PROCESS {
        try {
            if (-not $ExcludeSpecialCharacters) {
                   $Password = -join ((48..57) + (65..90) + (97..122) + (97..122) + $SpecialCharacters | Get-Random -Count $Length | foreach {[char]$_})
                }
            else {
                   $Password = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $Length | foreach {[char]$_})  
                }
        } 
        catch {
           throw $_.Exception.Message
        }
    }
 
    END {
        Write-Output "Here is your random password: $($Password)"
    }