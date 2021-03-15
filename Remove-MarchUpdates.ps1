
# Get the OS build number so we know which update package to remove
$BuildNumber = (Get-ComputerInfo).OSBuildNumber

# Email parameters
$MailRecipient = "<EMAIL_ADDRESS_OF_DISTRIBUTION_GROUP>"
$MailSender = "<INSERT_PROPER_ACTOR>-$env:COMPUTERNAME@<EMAIL_DOMAIN>"
$Subject = "[Success] <INSERT_PROPER_ACTOR> - March Update Removal"
$MailSMTPServer = "aspmx.l.google.com"

$LocalLogLocation = "C:\Windows\Temp\$env:COMPUTERNAME-march-update-removal-error.txt"



# If the build is 20H2 or 2004
if ($BuildNumber -eq '19042' -or $BuildNumber -eq '19041')
{
    try
    {
        # The update version for the "RollupFix" (does not show up as a KB in DISM)
        $UpdateVersion = "19041.867.1.8"
        # Select just the package name - no fluff
        $Update = "$(dism /online /get-packages | Select-String -Pattern "Package_for" | Select-String -Pattern "$UpdateVersion")".split(":")[1].replace(" ", "")

        if ( $Update )
        {
            # Invoke DISM to remove the package quietly and don't reboot
            dism /Online /Remove-Package /PackageName:$Update /quiet /norestart

            $Body = "Computer $env:COMPUTERNAME removed the update ' $Update ' from the system`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

            # Send the email notification
            Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $Body -SmtpServer $MailSMTPServer

            # Create a pop-up window to notify the user to reboot
            #$wsh = New-Object -ComObject Wscript.Shell
            #$wsh.Popup("An automated removal of the Microsoft March Update has been performed.`n`nTo complete the uninstallation (and be safe on printing), please reboot your machine ASAP.")
        }
        else
        {
            # The update was not found on the system - it must have been removed already
        }

    }

    catch
    {

        Write-Output "$($_ | Select-Object -Property *)" > $LocalLogLocation

        $Subject = "[Failure] <INSERT_PROPER_ACTOR> - March Update Removal"

        $Body = "SComputer $env:COMPUTERNAME attempted to remove an update but failed! The error message is attached.`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        # Send the email notification
        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $Body -Attachments $LocalLogLocation
    }
}

# If the build is 1909
elseif ($BuildNumber -eq '18363')
{
    try
    {
        # The update version for the "RollupFix" (does not show up as a KB in DISM)
        $UpdateVersion = "18362.1440.1.7"
        # Select just the package name - no fluff
        $Update = "$(dism /online /get-packages | Select-String -Pattern "Package_for" | Select-String -Pattern "$UpdateVersion")".split(":")[1].replace(" ", "")

        if ( $Updates )
        {
            # Invoke DISM to remove the package quietly and don't reboot
            dism /Online /Remove-Package /PackageName:$Update /quiet /norestart

            $Body = "Computer $env:COMPUTERNAME removed the update ' $Update ' from the system`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

            # Send the email notification
            Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $Body -SmtpServer $MailSMTPServer

            # Create a pop-up window to notify the user to reboot
            #$wsh = New-Object -ComObject Wscript.Shell
            #$wsh.Popup("An automated removal of the Microsoft March Update has been performed.`n`nTo complete the uninstallation (and be safe on printing), please reboot your machine ASAP.")
        }
        else
        {
            # The update was not found on the system - it must have been removed already
        }

    }

    catch
    {
        Write-Output "$($_ | Select-Object -Property *)" > $LocalLogLocation

        $Subject = "[Failure] <INSERT_PROPER_ACTOR> - March Update Removal"

        $Body = "SCCM client on $env:COMPUTERNAME removed attempted to remove an update but failed! The error message is attached.`n`n`n`I'm a bot and this message is automated. Replies will not be monitored. ;) glhf"

        # Send the email notification
        Send-MailMessage -To $MailRecipient -From $MailSender -Subject $Subject -Body $Body -Attachments $LocalLogLocation
    }
}

else
{
    # The build number is not supported at this time    
}