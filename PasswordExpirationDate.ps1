# RW Knight (20171121)
# Get Date and Time and put in Variable
$CurrentDate = Get-Date
$CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh-mm')

# Get Computername
$Computername = $env:COMPUTERNAME
$Computername = $Computername.ToLower()

# Get Username
$Username = $env:USERNAME
$Username = $Username.ToLower()

# Get Email Sender
$Sender = "$Computername@pletter.com"

# Constant Variables
$OutputFile =  "C:\users\$Username\Password Expiration Dates $CurrentDate.csv"

# Get data and save to CSV
Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", "AccountExpirationDate" | Select-Object -Property "Displayname", @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, "AccountExpirationDate" | sort-object ExpiryDate | export-csv $OutputFile

# Open CSV in Excel (only use if run from user's desktop)
#$ExcelFile = $OutputFile
#$Excel = New-Object -Com Excel.Application
#$Excel.Visible = $True 
#$Workbook = $Excel.Workbooks.Open($ExcelFile)

# Send Reoprt via Email

## Make the Report Pretty
$head = @"
<Title>Password Expiration Dates</Title>
<style>
Body {
font-family: "Tahoma", "Arial", "Helvetica", sans-serif;
background-color:#FFFFFF;
}
table {
border-collapse:collapse;
width:60%
}
td {
font-size:12pt;
border:1px #808080 solid;
padding:5px 5px 5px 5px;
}
th {
font-size:14pt;
text-align:left;
padding-top:5px;
padding-bottom:4px;
background-color:#808080;
color:#FFFFFF;
}
name tr{
color:#000000;
background-color:#FFFFFFFF;
}
</style>
"@

## Output file at console and create message body
Import-Csv "C:\users\$Username\Password Expiration Dates $CurrentDate.csv"
$HTML = Import-Csv "C:\users\$Username\Password Expiration Dates $CurrentDate.csv" | ConvertTo-Html -head $head -precontent "<h2>Password Expiration Dates</h2>" -postcontent "<h6>report run date & time $CurrentDate</h6>" | Out-String

##Send message
$email = @{
From = $Sender
To = "rknight@pletter.com"
Subject = "Password Expiration Dates"
SMTPServer = "intmail.trsecure.com"
Attachments = $OutputFile
BodyAsHTML = $True
Body = $HTML
}
send-mailmessage @email

# Cleanup
Remove-Item $OutputFile

