# Task Scheduler EWS Alert
Powershell script to check for Task Scheduler tasks status and send email alert via EWS if task failed

Add to Task Scheduler
Start Program Powershell
Argument: -File "C:\apps\task_fail_alert\script.ps1"

Needs Exchange Managed API 2.2 to run (C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll to run).
