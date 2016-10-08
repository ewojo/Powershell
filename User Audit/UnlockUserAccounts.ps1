###############################################
#
#     Unlock AD Users Account With Prompt
#     Made By Eric Wojtowicz
#     Date 3/24/2016
#      
#
###############################################
Import-Module ActiveDirectory

$USERNAME=Read-Host "Enter the account name to unlock"

Unlock-ADAccount -identity $USERNAME

Write-Host "Account is unlocked for $USERNAME"