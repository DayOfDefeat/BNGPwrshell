##################################################################################################################
# Please Configure the following variables....
$smtpServer="intmail.trsecure.com"
#$expireindays = 5
$from = "<Passwordnotice@pletter.com>"
#$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "c:\log\log.csv" # ie. c:\mylog.csv
#$testing = "Disabled" # Set to Disabled to Email Users
#$date = Get-Date -format ddMMyyyy
#
###################################################################################################################
# Send Email Message
Send-Mailmessage -smtpServer $smtpServer -from $from -to "ITServices@pletter.com" -subject "Weekly Password Notification log" -Attachments $logFile -priority High
    # End Send Message
    # End User Processing