Get-ADUser -LDAPFilter "(!physicalDeliveryOfficeName=*)" -searchBase "OU=Users,DC=Contoso,DC=com" `
    | Select Name