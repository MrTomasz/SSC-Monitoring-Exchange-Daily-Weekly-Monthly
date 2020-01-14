# SSC-Monitoring-Exchange-Daily-Weekly-Monthly
Best practices and a proposition of monitoring solution for Exchange from Shared Services Canada, kick started by Kwok-Fai Ha, Shared Services Canada.

# Important for WindowsTaskDefinition files

Note that we will have to register the running user by using either the Windows Scheduler interface, or in bulk with PowerShell
using Set-ScheduledTask

# How to load all XML Windows Task definition into Windows

```powershell

$XMLFile = Get-ChildItem <Path to the folder where your XML task definition are stored - here .\WindowsTasksDefinition>

Register-ScheduledTask -xml (Get-Content 'C:\PATH\TO\IMPORTED-FOLDER-PATH\TASK-INPORT-NAME.xml' | Out-String) -TaskName "TASK-IMPORT-NAME" -TaskPath "\TASK-PATH-TASKSCHEDULER\" -User COMPUTER-NAME\USER-NAME –Force
```