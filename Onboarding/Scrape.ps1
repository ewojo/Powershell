#######################
#
# Caution, Thumbprint is unique to each user.  If a user receives a new thumbprint or thumbprint is unavailable, please enter current thumbprint.
#
$cert = (dir Cert:\CurrentUser\My\<cert>)
#######################
$r = Invoke-WebRequest -Uri '<url>' -Certificate $cert
$r.parsedhtml.getelementsbytagname("TR") |

    ForEach-Object {

      
        ( $_.children  |
            Where-Object { $_.tagName -eq "td" -or 'TH' } |
            Select-Object -ExpandProperty innerText | foreach {'"'+$_+'"'} 

        )  -join ',' 
    } | Out-File -Encoding ascii 'C:\Users\location\scrape.csv
