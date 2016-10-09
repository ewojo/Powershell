#Unlock AD Users account with prompt
Import-Module ActiveDirectory

$USERNAME=Read-Host "Enter the account name to unlock"

Unlock-ADAccount -identity $USERNAME

Write-Host "Account is unlocked for $USERNAME"