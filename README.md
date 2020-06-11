# ScheduledTasksReport

PowerShell script to generate a CSV report of scheduled tasks (Windows). Uses the Get-ScheduledTask cmdlet.
The report is for documentation of tasks withs basic task properties.

## Prerequisites

PowerShell Version >= 5.1 required.
Tested with PowerShell 7.0 and 5.1 (Windows 10 Pro)

## Input

* $taskPath : task folder(s) to scan
* $outcsv: path of CSV file

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