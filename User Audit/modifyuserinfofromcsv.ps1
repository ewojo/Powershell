# Import active directory module for running AD cmdlets
Import-Module activedirectory

# Allow user to enter location of file
Write-Host "Please enter the location of your file: "
$CSVLocation = Read-Host

# Import CSV from location given
$ADUsers = Import-csv $CSVLocation

# Loop through each row containing user details in the CSV file
foreach ($User in $ADUsers)
{
    # Read user data from each field in each row and assign the data to a variable as below

    $Username  = $User.username
    $Password  = $User.password
    $Firstname = $User.firstname
    $Lastname  = $User.lastname
    $OU        = $User.ou # This field refers to the OU the user account is to be created in

    # Check to see if the user exists in AD
    if (Get-ADUser -F {SamAccountName -eq $Username})
    {
        # If user does not exist, notify and ask for creation
        Write-Warning "User does not exist. Do you want to create a new user?"
        $selection = Read-Host

        # If "y", create account
        if($selection -eq "y")
            {
                Add-ADUser `
                -sAMAccountName $Username `
                -UserPrincipalName $Username@fgdevopslab.local `
                -Name "$Firstname $Lastname" `
                -GivenName $Firstname `
                -Surname $Lastname `
                -Enabled $True `
                -DisplayName "$Firstname $Lastname" `
                -Path $OU `
                -AccountPassword (convertto-securestring $Password -AsPlainText -Force)
            }
        else
        {
            Write-Host "User not created."
        }
    }
    else
    {
        # Account will be modified with the info from the CSV file
        Set-ADUser `
            -sAMAccountName $Username `
            -UserPrincipalName $Username@fgdevopslab.local `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$Firstname $Lastname" `
            -Path $OU `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force)
    }
}