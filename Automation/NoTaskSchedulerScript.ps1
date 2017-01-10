#Run Powershell Script without task scheduler
while ($true)
{
     #Insert Powershell script you want to run
     Write-Host 'Script Ran @' (Get-Date);
     #Put (60* minutes)
     sleep -seconds (60)
}
