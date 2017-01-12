
############  Global config variables Start ##################
$networkPath = "#EnterPath"
$DestinationPath = "#EnterPath";
$emailSubject = "#EnterSpecificSubject"
$LogFile = "ExportOLEmailsLogUS.txt";
############  Global config variables End ##################


############   Functions Start ##################
function GetScriptPath
{     Split-Path $myInvocation.ScriptName 
}
function ShowError  ($msg){Write-Host "`r`n";Write-Host -ForegroundColor Red $msg; LogToFile $msg $LogFile; }
function ShowSuccess($msg){Write-Host "`r`n";Write-Host -ForegroundColor Green  "$msg";LogToFile $msg $LogFile; }
function ShowProgress($msg){Write-Host "`r`n";Write-Host -ForegroundColor Cyan  "$msg";LogToFile $msg $LogFile; }
function ShowInfo($msg){Write-Host "`r`n";Write-Host -ForegroundColor Yellow  "$msg";LogToFile $msg $LogFile; }
function LogToFile   ($msg, $ouputFile)
{	
	$msg |Out-File -Append -FilePath $ouputFile -ErrorAction:SilentlyContinue;
}
############   Functions End  ##################

#Getting directory path from where script is running
$CurrentDir = GetScriptPath
$LogFile = "$CurrentDir\$LogFile"

$s = Get-Date;
ShowProgress "Script started at $s";
#Add Interop Assembly 
Add-type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null 
 
#Type declaration for Outlook Enumerations, 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 
$olSaveType = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type] 
$olClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type] 
 
#Add Outlook Com Object, MAPI namespace, and set folder to the Inbox 
$outlook = New-Object -ComObject Outlook.Application 
$namespace = $outlook.GetNameSpace("MAPI") 
#$folder = $namespace.getDefaultFolder($olFolders::olFolderInBox) 
#Folder OTHER than Inbox
$folder = ($namespace.getDefaultFolder($olFolders::olFolderInBox).folders|WHERE {$_.Name -eq "Not Inbox"})
 
 $copiedEmailsCount =0;
#Iterate through each object in the chosen folder 
ShowInfo "$($folder.Items.Count) emails to process";

$emailCount = 1;
foreach ($email in $folder.Items) 
{ 
    ShowInfo "Processing email no: $emailCount"; 
    #Get email's subject and date 
    [string]$subject = $email.Subject 
    [string]$sentOn = $email.SentOn
            try
                {
                    $date = [System.DateTime]::Parse($sentOn);
                    $sentOn = $date.ToString("MMM-yyyy");
       }
    catch{}
    [string]$senderName  = $email.senderName
      
 
    $sentOn = $sentOn.Replace("/","-")
	$sentOn = $sentOn.Replace(":","-")
	
    #Strip subject and date of illegal characters, add .msg extension, and combine 
    #$fileName = Remove-InvalidFileNameChars -Name($sentOn + "-" + $subject + ".msg") 
	 $fileName = "#FileName.msg"  
 
    #Combine destination path with stripped file name 
    $dest = "$DestinationPath\$fileName"
	 #Test if object is a MailItem 
    if ($email.Class -eq $olClass::olMail) 
	{          
	    if($subject -eq $emailSubject)
		{
			try
			{
		  
				$existingFileCheck =$null;
			 	$existingFileCheck = Get-Item $dest -ErrorAction:SilentlyContinue;
			 	if(-not $existingFileCheck)
			 	{
				   	ShowProgress "Saving email with Sender [$senderName] & subject [$subject]";
					$email.SaveAs($dest, $olSaveType::olMSG)
					ShowSuccess "Email saved as $dest";
					$copiedEmailsCount++; 
				}
				else
				{
				   ShowSuccess  "Email already copied.";
		 		}
		       
			}
			catch
			{
				$ErrorMessage = $_.Exception.Message
				ShowError "Email could not be saved. $ErrorMessage";
				break;
			}
     	}         
    } 
	
	$emailCount++;
} 
 
 if($copiedEmailsCount -gt 0)
 {
    ShowInfo "$copiedEmailsCount files saved successfully matching subject [$subject]"
 	$files =  Get-ChildItem $DestinationPath -Filter "*.msg";
	$copiedOverNetwork =0;
	foreach($file in $files)
	{
	    $fileToCopy= $file.FullName;
		$dest = "$networkPath\$file"	
		$existingFileCheck =$null;
		$existingFileCheck = Get-Item $dest -ErrorAction:SilentlyContinue;
		if(-not $existingFileCheck)
		{
		 	try
		 	{
		   		ShowProgress "Copying $fileToCopy";
		   		Copy-Item $fileToCopy -Destination $dest 
				ShowSuccess "Email copied as $dest";
				$copiedOverNetwork++;			
				
		   	}
			catch
			{
				$ErrorMessage = $_.Exception.Message
				ShowError "Email could not be copied. $ErrorMessage";			
			}
		 }
		 else
		 {
		  	ShowSuccess  "Email [$dest] already copied.";
		 }
	}
	
	if($copiedOverNetwork -gt 0)
	{
		ShowInfo "$copiedOverNetwork emails saved successfully to network path $networkPath"
	}
 }
 else
 {
 	ShowProgress "No email to copy to $networkPath";
 }
 $e = Get-Date;
ShowProgress "Script ended at $e";
