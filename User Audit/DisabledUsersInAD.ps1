###############################################
#
#     Find Disabled Users In OU
#     Made By Eric Wojtowicz
#     Date 3/24/2016
#      
#
###############################################
import-module activedirectory
Get-ADUser -Filter {Enabled -eq $false} -SearchBase “OU=DevOps,DC=Contoso,DC=com” -Properties * | 
    select -Property Name,DisplayName,mail,SamAccountName, Enabled |
    Sort-Object -Property SamAccountName | 
    Export-csv "C:\TEST\DisabledUsers.csv" 