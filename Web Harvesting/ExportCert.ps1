#export certificate from current computer
$cert = (Get-ChildItem -Path Cert:\CurrentUser\My\)
 Export-Certificate -Cert $cert -FilePath C:\Users\(*Insert Location Here\Cert.cer
