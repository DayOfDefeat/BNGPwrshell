<#	
	.NOTES
	===========================================================================
	 Created with: 	Blood, Sweat and Tears
	 Created on:   	11/20/2015 1:36 PM
	 Created by:   	Joe and Tyler
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		Automation script to run for users not logged in for 90 days - Mailboxes.
#>
Import-Module MSOnline
$username = "***"
$password = cat C:\random\txt.txt | convertto-securestring -AsPlainText -Force
$cred = new-object -typename System.Management.Automation.PSCredential `
				   -argumentlist $username, $password
Connect-MsolService –Credential $cred
Try
{
	$ErrorActionPreference = "Stop";
	if ((Get-MsolDomain) -ne $null)
	{
		Write-Host "SUCCESS: Connected to Microsoft Online " -ForegroundColor Green
	}
}
Catch
{
	Write-Host "FAILED: Unable to Connect to Microsoft Online" -ForegroundColor Red
}
Try
{
	$ErrorActionPreference = "Stop";
	$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
	Import-PSSession $exchangeSession -DisableNameChecking
	
	if ((Get-AcceptedDomain) -ne $null)
	{
		Write-Host "SUCCESS: Connected to Exchange Online " -ForegroundColor Green
	}
}
Catch
{
	Write-Host "FAILED: Unable to Connect to Exchange Online" -ForegroundColor Red
}
#$getDate = Get-Date
#$date = $getdate.Adddays(-90)
#Get-mailbox -resultsize unlimited | Get-MailboxStatistics | select displayname, lastlogontime | Where-Object {$_.lastlogontime -le $date} | Export-Csv C:\Random\test.csv

Get-mailbox -resultsize unlimited | Get-MailboxStatistics| where { $_.LastLogonTime -lt (get-date).AddDays(-90) } | select displayName, lastlogontime | Export-Csv C:\Random\test.csv

$From = "Noreply@***.com"
$To = "Servicedesk@***.com"
$Attachment = "C:\Random\test.csv"
$Subject = "Test"
$Body = "this is a test"
$SMTPServer = "<.>"
$SMTPPort = "25"
$RSusername = "Noreply@***.com"
$RSpassword = "***" | convertto-securestring -AsPlainText -Force
$win = new-object -typename System.Management.Automation.PSCredential -argumentlist $RSusername, $RSpassword

Send-MailMessage -From $From -to $To -Subject $Subject `
				 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
				 -Credential $win -Attachments $Attachment

##Certificate mis match therefore it breaks##

#$RSusername = "noreply@***.com"
#$RSpassword = cat C:\random\email.txt | convertto-securestring -AsPlainText -Force
#$RScred = new-object -typename System.Management.Automation.PSCredential -argumentlist $RSusername, $RSpassword

#send-mailmessage -from "noreply@***.com" -to "Servicedesk@***.com" -subject "Email last logon time audit report" -body "Please verify" -Attachments "C:\Random\test.csv" -smtpserver smtp.emailsrvr.com -Port 25 -Credential $RScred

$Session = Get-PSSession
Remove-PSSession -id $Session.Id