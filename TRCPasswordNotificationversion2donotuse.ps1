#################################################################################################################
# 
# Version 1.3 April 2015
# Robert Pearman (WSSMB MVP)
# TitleRequired.com
# Script to Automated Email Reminders when Users Passwords due to Expire.
#
# Requires: Windows PowerShell Module for Active Directory
#
# For assistance and ideas, visit the TechNet Gallery Q&A Page. http://gallery.technet.microsoft.com/Password-Expiry-Email-177c3e27/view/Discussions#content
#
##################################################################################################################
# Please Configure the following variables....
$smtpServer="Mail1.tresearch.net"
$expireindays = 5
$from = "NoReply<Noreply@pletter.com>"
$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "c:\log\log.csv" # ie. c:\mylog.csv
$testing = "Enabled" # Set to Disabled to Email Users
$testRecipient = "Jhuiet@pletter.com"
$date = Get-Date -format ddMMyyyy
# Define font and font size
# ` or \ is an escape character in powershell
$font = "<font size=`"3`" face=`"Calibri`">"
#
#Generate a Admin report?
$ReportToAdmin = $true
#$ReportToAdmin = $false
#
#Sort Report
#===========================#
# 0 = By OU
# 1 = First Name Ascending
# 2 = Last Name Ascending
# 3 = Expiration Date Ascending
# 4 = First Name Descending
# 5 = Last Name Descending
# 6 = Expiration Date Descending
#===========================#
$ReportSortBy=1
###################################################################################################################
# Sort Report
Switch ($ReportSortBy)
{
  '0' {$users}
  '1' {$users = $users | sort {$_.FirstName}}
  '2' {$users = $users | sort {$_.LastName}}
  '3' {$users = $users | sort {$_.PasswordExpires}}
  '4' {$users = $users | sort -Descending {$_.FirstName}}
  '5' {$users = $users | sort -Descending {$_.LastName}}
  '6' {$users = $users | sort -Descending {$_.PasswordExpires}}
}

if ($ReportToAdmin -eq $true)
{
  #Headings used in the Admin Alert
  $Title_ExpiredNoEmail="<h3><u>Expired users with no email address</h3></u>"
  $Title_AboutToExpireNoEmail="<h3><u>Users about to expire with no email address</h3></u>"
  $Title2="<br><br><h2><u><font color= red>No Admin Action Required - Email Sent to User</h2></u></font>"
  $Title_Expired="<h3><u>Expired users that have been notified</h3></u>"
  $Title_AboutToExpire="<h3><u>Users about to expire that have been notified</h3></u>"
  $Title_NoExpireDate="<h3><u>Users with no expiration date</u></h3>"
}
# Check Logging Settings
if (($logging) -eq "Enabled")
{
    # Test Log File Path
    $logfilePath = (Test-Path $logFile)
    if (($logFilePath) -ne "True")
    {
        # Create CSV File and Headers
        New-Item $logfile -ItemType File
        Add-Content $logfile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn"
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
    <p> This is a auto-generated email to remind you that your TRC Network Password will expire $messageDays.<br>   
    Please reset it now by pressing Ctrl+Alt+Del on your keyboard and choosing to Reset Password. Please also connect to https://mail.pletter.com and change your Outlook password.<br>
	If you are working remotely, please ensure that you connect to the VPN before you reset your password.<br> 
	Please reach out to ITservices should you need assistance at ITservices@pletter.com<br> 
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
        Send-Mailmessage -smtpServer $smtpServer -from $from -to "Jhuiet@pletter.com" -Cc "" -subject $subject -body $body -bodyasHTML -priority High  

    } # End Send Message
    #For loop to report
foreach ($user in $users)
{

  if ($user.PasswordExpires -eq $null)
  {
    $UsersList_WithNoExpiration += $user.Name + " (<font color=blue>" + $user.LogonName + "</font>) does not seem to have a Expiration Date on their account.<font color=Green> <br>OU Container: " + $user.DN + "</font> <br>"
  }
  Elseif ($user.PasswordExpires -ne $null)
  {
    #Calculate remaining days till Password Expires
    #$DaysLeft = (($user.PasswordExpires - $date).Days)

    #Days till password expires
    #$DaysLeftTillExpire = [Math]::Abs($DaysLeft)

    #If password has expired
    If ($DaysLeft -le 0)
    {
      #If the users don't have a primary SMTP address we'll report the problem in the Admin report
      if (($user.Email -eq $null) -and ($user.UserMustChangePassword -ne $true) -and ($ReportToAdmin -eq $true))
      {
        #Add it to admin list to report on it
        $UserList_ExpiredNoEmail += $user.name + " (<font color=blue>" + $user.LogonName + "</font>) password has expired " + $DaysLeftTillExpire + " day(s) ago</font>. <font color=Green> <br>OU Container: " + $user.DN + "</font> <br><br>"
      }

      #Else they have an email address and we'll add this to the admin report and email the user.
      elseif (($user.Email -ne $null) -and ($user.UserMustChangePassword -ne $true) -and ($AlertUser -eq $true))
      {
        #$ToAddress = Jhuiet@pletter.com
       # $Subject = "Friendly Reminder: Your TRC Network Password has expired."
       # $body = " "
       # $body = $font
       # $body += "Hello " + $user.Name + ",<br><br>"
       # $body += "This is a auto-generated email to remind you that your TRC Network Password for account - <font color=red>" + $user.LogonName + "</font> - has expired. <br><br>"
       # $body += "Please contact IT for assistance."
       # $body += "<br><br><br><br>"
      #  $body += " "
       # $body += "<h4>Note: Never Share Your Password With Others!</h4>"
       # $body += "</font>"

       # Send-MailMessage -smtpServer $RelayMailServer -from $FromAddress -to $user.Email -subject $Subject -BodyAsHtml -body $body

      }
      if ($ReportToAdmin -eq $true)
      {
        #Add it to a list
        $UserList_ExpiredHasEmail += $user.name + " (<font color=blue>" + $user.LogonName + "</font>) password has expired " + $DaysLeftTillExpire + " day(s) ago</font>. <font color=Green> <br>OU Container: " + $user.DN + "</font> <br><br>"
      }
    }
    elseif ($DaysLeft -ge 0)
    {
      #If Password is about to expire but the user doesn't have a primary address report that in the Admin report
      if (($user.Email -eq $null) -and ($user.UserMustChangePassword -ne $true) -and ($ReportToAdmin -eq $true))
      {
        #Add it to admin list
        $UserList_AboutToExpireNoEmail += $user.name + " (<font color=blue>" + $user.LogonName + "</font>) password is about to expire and has " + $DaysLeftTillExpire + " day(s) left</font>. <font color=Green> <br>OU Container: " + $user.DN + "</font> <br><br>"
      }
      # If there is an email address assigned to the AD Account send them a email and also report it in the admin report
      elseif (($user.Email -ne $null) -and ($user.UserMustChangePassword -ne $true) -and ($AlertUser -eq $True) )
      {
        #Setup email to be sent to user
      #  $ToAddress = $user.Email
      #  $Subject = "Notice: Your TRC Network Password is about to expire."
      #  $body = " "
      #  $body = $font
      #  $body += "Hello " + $user.Name + ",<br><br>"
      #  $body += "This is a auto-generated email to remind you that your TRC Network Password for account - <font color=red>" + $user.LogonName + "</font> - will expire in </font color = red>" + $DaysLeftTillExpire +" Day(s). <br><br>"
      #  $body += "Please reset it now by pressing Ctrl+Alt+Del on your keyboard and choosing to Reset Password.  "
      #  $body += "If you are working remotely, please make sure that you have an active VPN connection when you reset it."
      #  $body += "<br><br><br><br>"
        #$body += "<h4>Note: Never Share Your Password With Others!</h4>"
       # $body += "</font>"

        #Send-MailMessage -smtpServer $RelayMailServer -from $FromAddress -to $user.Email -subject $Subject -BodyAsHtml -body $body

      }
      if ($ReportToAdmin -eq $true)
      {
        #Add it to admin Report list
        $UserList_AboutToExpire += $user.name + "&nbsp; <font color=blue>(" + $user.LogonName + "</font>) password is about to expire and has " + $DaysLeftTillExpire + " day(s) left</font>. <font color=Green> <br>OU Container: " + $user.DN + "</font> <br><br>"
      }
    }
  }
} # End foreach ($user in $users)

if ($ReportToAdmin -eq $true)
{
  If ($UserList_AboutToExpire -eq $null) {$UserList_AboutToExpire = "No Users to Report"}
  If ($UserList_AboutToExpireNoEmail -eq $null){ $UserList_AboutToExpireNoEmail = "No Users to Report"}
  if ($UserList_ExpiredHasEmail -eq $null) {$UserList_ExpiredHasEmail = "No Users to Report"}
  if ($UserList_ExpiredNoEmail -eq $null) {$UserList_ExpiredNoEmail = "No Users to Report"}
  if ($UsersList_WithNoExpiration -eq $null) {$UsersList_WithNoExpiration = "No Users to Report"}

  #Email Report to Admin
  $Subject="Password Expiration Status for " + $today + "."
  $Footer = "<br /><br /><p><font color=#666666 size=2><strong>Note:</strong> This script was run from $($ENV:ComputerName) via scheduled task.</font></p>"
  $AdminReport = $font + $Title + $Title_ExpiredNoEmail + $UserList_ExpiredNoEmail + $Title_AboutToExpireNoEmail + $UserList_AboutToExpireNoEmail + $Title_AboutToExpire + $UserList_AboutToExpire + $Title_Expired + $UserList_ExpiredHasEmail + $Title_NoExpireDate + $UsersList_WithNoExpiration + "</font>" + $Footer
  Send-MailMessage -smtpServer $RelayMailServer -from $FromAddress -to "Jhuiet@pletter.com" -subject $Subject -BodyAsHtml -body $AdminReport

     #Email Report to Admin
 # $Subject="Password Expiration Status for " + $today + "."
  #$Footer = "<br /><br /><p><font color=#666666 size=2><strong>Note:</strong> This script was run from $($ENV:ComputerName) via scheduled task.</font></p>"
  #$AdminReport = $Name + "</font>" + $Footer
  #Send-MailMessage -smtpServer $SmtpServer -from $from -to "Jhuiet@pletter.com" -subject $Subject -BodyAsHtml -body $AdminReport
} # End User Processing

}

# End