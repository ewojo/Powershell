#$Cert = Import Certificate Location if applicable 
$r = Invoke-WebRequest -Uri '#https://website' #-Certificate $Cert
$r.parsedhtml.getelementsbytagname("TR") |

    ForEach-Object {

      
        ( $_.children  |
            Where-Object { $_.tagName -eq "td" -or 'TH' } |
            Select-Object -ExpandProperty innerText | foreach {'"'+$_+'"'} 

        )  -join ',' 
    } | Out-File -Encoding ascii "C:\Users\ewojo\Desktop\users.csv"
