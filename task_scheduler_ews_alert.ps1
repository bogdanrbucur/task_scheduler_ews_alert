#Add to Task Scheduler
#Start Program Powershell
#Argument: -File "C:\apps\task_fail_alert\script.ps1"

#Web Service Path
param(
$mailbox="aceapps@ace-tankers.com",
$userName="acetankers\aceapps",
$password="Queen!@2017"
)

#Importing WebService DLL
$EWSServicePath="C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $EWSServicePath 

#Creating Service Object
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$EWSurl = "https://webmail.ace-tankers.com/ews/Exchange.asmx"
$Service.URL = $EWSurl

#Setting up ImperSonated User - only useful if not running as mailbox user
# $ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$mailbox
# $ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
# $service.ImpersonatedUserId = $ImpUserId


# Not all codes >0 are fail codes... must exclude "running" status codes

# Tasks to be checked for. Rename to TaskName2 etc. for additional tasks
# Task name must match the Name from Task Scheduler

# QA Forms Analyzer
$ScheduledTaskName1 = 'QA Forms Analyzer'
$TaskResult1 = (Get-ScheduledTaskInfo -TaskName $ScheduledTaskName1).LastTaskResult
If ($TaskResult1 -gt 0) {

	#Setting up Email message Class
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
	#Creating and Saving Message in Draft Folder
	$message.Subject = "SPDEV task $ScheduledTaskName1 failed with code $TaskResult1"
	$message.From = $mailbox
	$message.ToRecipients.Add("bogdanb-it@asm-maritime.com")
	$message.ToRecipients.Add("naved-it@asm-maritime.com")
	$message.Body = "Login as aceapps to 192.168.222.158 to check the status."
	$message.SendAndSaveCopy()
	}
	
# QA Approval
$ScheduledTaskName2 = 'QA Approval'
$TaskResult2 = (Get-ScheduledTaskInfo -TaskName $ScheduledTaskName2).LastTaskResult
If ($TaskResult2 -gt 0) {

	#Setting up Email message Class
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
	#Creating and Saving Message in Draft Folder
	$message.Subject = "SPDEV task $ScheduledTaskName2 failed with code $TaskResult2"
	$message.From = $mailbox
	$message.ToRecipients.Add("bogdanb-it@asm-maritime.com")
	$message.ToRecipients.Add("naved-it@asm-maritime.com")
	$message.Body = "Login as aceapps to 192.168.222.158 to check the status."
	$message.SendAndSaveCopy()
	}
	
# TelexINOut
$ScheduledTaskName3 = 'TelexINOut'
$TaskResult3 = (Get-ScheduledTaskInfo -TaskName $ScheduledTaskName3).LastTaskResult
If ($TaskResult3 -gt 0) {

	#Setting up Email message Class
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
	#Creating and Saving Message in Draft Folder
	$message.Subject = "SPDEV task $ScheduledTaskName3 failed with code $TaskResult3"
	$message.From = $mailbox
	$message.ToRecipients.Add("bogdanb-it@asm-maritime.com")
	$message.ToRecipients.Add("naved-it@asm-maritime.com")
	$message.Body = "Login as aceapps to 192.168.222.158 to check the status."
	$message.SendAndSaveCopy()
	}
	
# PAL Exchange Rates Interface
$ScheduledTaskName4 = 'PAL-ExchangeRatesInterface'
$TaskResult4 = (Get-ScheduledTaskInfo -TaskName $ScheduledTaskName4).LastTaskResult
If ($TaskResult4 -gt 0) {

	#Setting up Email message Class
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
	#Creating and Saving Message in Draft Folder
	$message.Subject = "SPDEV task $ScheduledTaskName4 failed with code $TaskResult4"
	$message.From = $mailbox
	$message.ToRecipients.Add("bogdanb-it@asm-maritime.com")
	$message.ToRecipients.Add("naved-it@asm-maritime.com")
	$message.Body = "Login as aceapps to 192.168.222.158 to check the status."
	$message.SendAndSaveCopy()
	}
exit