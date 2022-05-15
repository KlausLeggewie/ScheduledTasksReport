# ScheduledTasksReport

PowerShell script to generate an Excel report of scheduled tasks (Windows). 
The report is for documentation of tasks with basic task properties.


Uses
* Get-ScheduledTask cmdlet
* ImportExcel Package from PowerShell Gallery

## Prerequisites

* PowerShell Version >= 5.1 required
* ImportExcel Package installed (https://www.powershellgallery.com/packages/ImportExcel)

Tested with PowerShell 7.0 and 5.1 (Windows 10 Pro)

## Customizable Parameters

* $taskPath : task folder(s) to scan
* $outXlsxPath: path of Excel file

## Output

* Name
* Path
* Description
* Status
* Command
* Arguments
* ScheduleEnabled
* StartDate
* EndDate
* StartTime
* Interval
* DaysOfWeek
* Months
* DaysOfMonth
* Repetition
* Duration
  
Not all possible definitions of a task are covered, e. g. if there are multiple triggers per task.

## Known Issues

The Get-ScheduledTask cmdlet does not fetch MonthlyTriggers correctly. Therefore monthly triggers are reported as interval "Other" (instead of "Monthly").