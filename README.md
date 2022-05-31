# Task Scheduler EWS Alert
Powershell script to check for Task Scheduler tasks status and send email alert via EWS if task failed

Add to Task Scheduler
Start Program Powershell
Argument: -File "C:\apps\task_fail_alert\script.ps1"

Needs Microsoft.Exchange.WebServices.dll to run (Exchange Managed API 2.2).
