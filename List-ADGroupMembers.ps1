<#  
.SYNOPSIS
    This function takes an AD Group name from the text box and prints the group members on the screen.

.DESCRIPTON
    Use this script for users who want to know AD group membership.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> List-AdGroupMembership.ps1 -ADGroupName "GroupName"

.PARAMETER ADGroupName
    This is the name of the AD Group. It takes a string data type.

.NOTES
    Created by: Josh Crabtree 26 March 20
#>

[cmdletbinding()]

param(
    [Parameter (
        Mandatory = $true,
        HelpMessage = "Enter the name of an Active Directory Group below and then click the submit button. The users in that group will appear on the screen."
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADGroupName
)

begin{
    $results = @()
}

process{
    #use 'Get-AdGroupMember' to get actual members of AD Group a user types
    $groupmembers = Get-ADGroupMember -Identity $ADGroupName -Recursive
    #return a count of the total number of people in the group
    Write-Output "Total Members in Group:"$groupmembers.count
    $results += foreach($user in $groupmembers){get-aduser $user.SamAccountName -Properties *}
}

end{
    $results|ft name,mail,department
}