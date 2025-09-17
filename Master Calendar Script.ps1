#---------------------------------------------------------------------
# Purpose: Multi calendar script   | Version: 1.4                    |
# Author:  Kye Hodgson             | Made for: Inreach Group         |
# Date:    31/03/2025              |                                 |
#---------------------------------------------------------------------

#------------------- Group to room calendar permisssions -------------
$GURCP = {
# Define the distribution group and mailbox
$distGroup = Read-Host "what is the distribution group you are giving access?"
$mailbox = Read-Host "What is the mail box you are giving access to?"
$mailboxcal = $mailbox+":\Calendar"
Write-Host $mailboxcal
# Get all the users in the distribution group
$users = Get-DistributionGroupMember -Identity $distGroup
Write-Host $user.PrimarySmtpAddress

# Loop through each user and add them to the mailbox calendar with editor access
foreach ($user in $users) {
    Add-MailboxFolderPermission -Identity $mailboxcal -User $user.PrimarySmtpAddress -AccessRights Editor
    }
Get-MailboxFolderPermission -Identity $mailboxcal
    $ask = Read-Host "Again? 1 Yes or 0 No"

If ($ask -eq 1){
    &$GURCP 
    }
elseif($ask-eq 0){
    Disconnect-ExchangeOnline
    exit
    }
}

#------------------- User to user calendar perms ---------------------
$USerCalendarChanger = {

$cal = ":\calendar"

Write-Host "what is the user whos calendar needs to be accessed? i.e KyeH"
$firstusername = Read-Host "First User Name"
$auser1 = $firstusername+$cal

Write-Host "who is the user you need to give access to? i.e JohnD"
$secondusername = Read-Host "Second User Name"

Write-Host "What permisssions does the second user need on the first users calendar"
Write-Host "Author, Owner, Editor, Reviewer, None?"
$accessR = Read-Host "Access Level"

Get-MailboxFolderPermission -Identity $auser1
Add-MailboxFolderPermission -Identity $auser1 -user $secondusername -AccessRights $accessR
pause
$ask = Read-Host "Again? 1 for Yes or 0 for No"

If ($ask -eq 1){
    &$USerCalendarChanger
    }
elseif($ask -eq 0){
    Disconnect-ExchangeOnline
    exit
    }
}


#-----------------user to multiple user calendar
$userToMulti = {

$UserAccess = Read-host "Who is the user you are giving everyone access rights to?"
$userAccessRights = Read-host "what is the access rights you are giving?"

$users = ﻿Get-Mailbox | Select -ExpandProperty Alias

Foreach ($user in $users) {Add-MailboxFolderPermission $user":\Calendar" -user $UserAccess -accessrights $userAccessRight}

Get-MailboxFolderPermission -Identity $mailboxcal
    $ask = Read-Host "Again? 1 Yes or 0 No"

If ($ask -eq 1){
    &$userToMilti 
    }
elseif($ask-eq 0){
    Disconnect-ExchangeOnline
    exit
    }
}
#-----------------Basic Logic Table-------------

$logicTable = {
$message1 = "What Account is this for?"
$question1 = "1) User to user, 2) Group to mailbox 3) exit"

$choices1 = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices1.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 User to a single User mailbox calendar'))
$choices1.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Group to a single Mailbox calendar'))
$choices1.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Provide everyone access to a single mailbox calendar'))
$choices1.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Exit'))

$decision1 = $Host.UI.PromptForChoice($message1, $question1, $choices1, 0)

If($decision1 -eq 0){
    &$USerCalendarChanger
}
elseif($decision1 -eq 1){
    &$GURCP
}
elseif($decision1 -eq 3){
    &$userToMulti
}
elseif($decision1 -eq 4){
    Disconnect-ExchangeOnline
    exit
}

}

#---------------Start Point-------------------
Connect-ExchangeOnline
&$logicTable