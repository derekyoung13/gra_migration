$action = New-ScheduledTaskAction -Execute “powershell.exe” -Argument "-executionpolicy bypass -file c:\programdata\GRA_O365\migration.ps1"
$trigger = New-ScheduledTaskTrigger -Once -At "3/9/2022 10:00:00 AM"
$principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -expand UserName)
$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
Register-ScheduledTask GRA_O365_Migration -InputObject $task
