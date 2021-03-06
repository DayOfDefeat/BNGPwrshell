#################################################################################################################
##################################################################################################################
# Please Configure the following variables....
$smtpServer="Intmail.trsecure.com"
$expireindays = 5
$from = "<ITDesk@pletter.com>"
$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "c:\log\log.csv" # ie. c:\mylog.csv
$testing = "Enabled" # Set to Disabled to Email Users
$testRecipient = "Jhuiet@pletter.com"
$date = Get-Date -format ddMMyyyy
#
###################################################################################################################

# Check Logging Settings
if (($logging) -eq "Enabled")
{
    # Test Log File Path
   # $logfilePath = (Test-Path $logFile)
    #if (($logFilePath) -ne "True")
    {
        # Create CSV File and Headers
        New-Item $logfile -ItemType File
        Out-File -Append $logfile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn" 
    }
} # End Logging Check

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
Import-Module ActiveDirectory
$users = get-aduser -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false }
$DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Process Each User for Password Expiry
foreach ($user in $users)
{
    $Name = $user.Name
    $emailaddress = $user.emailaddress
    $passwordSetDate = $user.PasswordLastSet
    $PasswordPol = (Get-AduserResultantPasswordPolicy $user)
    # Check for Fine Grained Password
    if (($PasswordPol) -ne $null)
    {
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge
    }
    else
    {
        # No FGP set to Domain Default
        $maxPasswordAge = $DefaultmaxPasswordAge
    }

  
    $expireson = $passwordsetdate + $maxPasswordAge
    $today = (get-date)
    $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
        
    # Set Greeting based on Number of Days to Expiry.

    # Check Number of Days to Expiry
    $messageDays = $daystoexpire

    if (($messageDays) -ge "1")
    {
        $messageDays = "in " + "$daystoexpire" + " days."
    }
    else
    {
        $messageDays = "today."
    }

    # Email Subject Set Here
    $subject="Your password will expire $messageDays"
  
    # Email Body Set Here, Note You can use HTML, including Images.
    $body ="
    Hello $name,
    <p> This is an email from IT Services to remind you that your TRC Network Password will expire $messageDays<br><br>   
    
	<b>Reset your Network (Computer) password from the Stockton Office:</b><br>
	1.  Press Ctrl+Alt+Del on your keyboard and choose Change a Password.<br><br>
	
	<b>Reset your Network Password Remote - Remote Users thru NetExtender (Important: Please complete all 5 steps in the order listed)</b><br>
    1.	Log into NetExtender with your current password.<br>
    2.	Press Ctrl+Alt+Del on your keyboard and choose Change a password.<br>
	3.  Enter old password then new password and confirm it.<br>
    4.	<b>While still connected to NetExtender</b>, press Ctrl+Alt+Del again on the keyboard and press Enter to lock the computer.<br>
    5.	Unlock your computer by pressing Ctrl+Alt+Del and log in with the new password.<br>
    <br> 
	
	<b>Reset your Outlook password:</b><br>
	1.  Go to http://mail.pletter.com.<br>
	2.  Click on the gear icon in the upper right corner of the menu bar.<br>
	3.  If you do not see a change password link, type password in the Search bar.<br>
	4.  Click the Password link to change your password.<br>
	5.	Open Outlook, you will be prompted to enter that new password into an Outlook login box.<br>
     <br>
	If you need assistance, or your password has already expired please contact ITDESK@pletter.com.<br> 
    <p>Thank you,<br> 
    <br>
    <p>IT Services<br>
    </P>"

   
    # If Testing Is Enabled - Email Administrator
    if (($testing) -eq "Enabled")
    {
        $emailaddress = $testRecipient
    } # End Testing

    # If a user has no email address listed
    if (($emailaddress) -eq $null)
    {
        $emailaddress = $testRecipient    
    }# End No Valid Email

    # Send Email Message
    if (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays))
    {
         # If Logging is Enabled Log Details
        if (($logging) -eq "Enabled")
        {
            Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson" 
        }
        # Send Email Message
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High  
        #
		#
	} # End Send Message
		

        #Send-Mailmessage -smtpServer $smtpServer -from $from -to "Jhuiet@pletter.com" -subject "Weekly Password Notification log" -Attachments $logfile -priority High
    
    
} # End User Processing



# End