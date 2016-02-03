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
$smtpServer="..."
$expireindays = 14
$from = "Administrator ..." # ie. admin@admin.net
$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "...\mylog.csv" # ie. c:\mylog.csv
$logFile1 = "...\mylog1.csv" # ie. c:\mylog1.csv
$testing = "Disabled" # Set to Disabled to Email Users
$testRecipient = "..."
$date = Get-Date -format ddMMyyyy
#
###################################################################################################################

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
$users = get-aduser -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } #| where { $_.passwordexpired -eq $false }
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
        $messageDays = "$daystoexpire" + " дней"
    }
    else
    {
        $messageDays = "сегодня"
    }

    # Email Subject Set Here
    $subject="Срок действия вашего пароля истекает через $messageDays"
  
    # Email Body Set Here, Note You can use HTML, including Images.
    $body ="
    Уважаемый пользователь $name,
    <p> Срок действия вашего пароля истекает через $messageDays.<br>
    Для изменения вашего пароля на компьютере нажмите CTRL ALT Delete и выберети меню смена пароля<br>
    <p>Спасибо<br> 
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
        $encoding = [System.Text.Encoding]::UTF8
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $encoding
     
    } # End Send Message
    
    # Send Email Message when password already expire
    if ($daystoexpire -lt "0")
    {
         $subject="Срок действия вашего пароля истёк"
         $body ="
                Уважаемый пользователь $name,
                <p> Срок действия вашего пароля истек.<br>
                Для изменения вашего пароля на компьютере нажмите CTRL ALT Delete и выберети меню смена пароля<br>
                <p>Спасибо<br> 
                </P>"
        # If Logging is Enabled Log Details
        if (($logging) -eq "Enabled")
        {
            Add-Content $logfile1 "$date,$Name,$emailaddress,$daystoExpire,$expireson"
        }
        # Send Email Message
        $encoding = [System.Text.Encoding]::UTF8
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $encoding
    } # End Send Message
    
} # End User Processing

# End
