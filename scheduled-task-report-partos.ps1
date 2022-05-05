# Create CSV Report of scheduled Tasks (for Windows)
# -------------------------------------------
# take care of PS version. $PSVersionTable.PSVersion
# this script is done for Version >= 5.1
# basic example: https://devblogs.microsoft.com/scripting/weekend-scripter-use-powershell-to-document-scheduled-tasks/

# pattern for task to fetch
$taskPath = "\Microsoft\*"
#name of the output file
$outcsv = "c:\temp\My Scheduled Tasks-$(Get-Date -Format "yyy-MM-dd").csv"


[Flags()] enum WeekDayFlag
{
Sun = 1
Mon = 2
Tue = 4
Wed = 8
Thu = 16
Fri = 32
Sat = 64
}

# reduce filepath to filename
function GetFileNameFromPath($path)
{
    if ($path -Match "\\") {
        try {
            return [io.path]::GetFileName($path)
        }
        catch {
            Write-Host ("GetFileName failed for ",$path , "; return empty filename")
            return ""
        }
    }
    return $path
}


Get-ScheduledTask -TaskPath $taskPath |
# optional: use "-TaskName" for specific task

    ForEach-Object { 
        [pscustomobject]@{

        Name = $_.TaskName
        Path = $_.TaskPath
        Description = $_.Description

        # if needed:
            #LastResult = $(($_ | Get-ScheduledTaskInfo).LastTaskResult)
            #NextRun = $(($_ | Get-ScheduledTaskInfo).NextRunTime)

        Status = $_.State 

        #reduce exe-path to filename
        Command = GetFileNameFromPath($_.Actions.execute)
        Arguments = $_.Actions.Arguments 

        ScheduleEnabled =   if (($_.Triggers.Enabled) -ne $null)
                                {($_.Triggers.Enabled | Select-Object -first 1)}
                            else 
                                {"N/A"} 
        
        StartDate = if (($_.Triggers.count -gt 0) -and ($_.Triggers[0].StartBoundary -ne $null)) 
                        {(Get-Date -Date ($_.Triggers[0].StartBoundary) -Format "d")}
                    else 
                        {"N/A"}
          
        EndDate =   if (($_.Triggers.EndBoundary) -ne $null) 
                        {($_.Triggers.EndBoundary)}
                    else 
                        {"N/A"} 

        StartTime = if (($_.Triggers.count -gt 0) -and ($_.Triggers[0].StartBoundary -ne $null)) 
                        {(Get-Date -Date ($_.Triggers[0].StartBoundary) -Format "HH:mm:ss")}
                    else 
                        {"N/A"}

        Interval =  if ($_.Triggers.count -eq 0) 
                        {"On Demand"}
                    elseif ($_.Triggers.DaysInterval -eq 1) 
                        {"Daily"}
                    elseif ($_.Triggers.WeeksInterval -eq 1)  
                        {"Weekly"}
                    elseif ($_.Triggers.ScheduleByMonth -eq 1) 
                        {"Monthly"}
                    else 
                        {"Other"}
          
        # Triggers.DaysOfWeek is a bitwise mask
        DaysOfWeek =    if (($_.Triggers.count -gt 0) -and ($_.Triggers[0].DaysOfWeek -ne $null))
                            {[WeekDayFlag]($_.Triggers[0].DaysOfWeek)}
                        else
                            {"N/A"} 

        #remark: output for month does not work, bug of the Get-ScheduledTask cmdlet
        #https://docs.microsoft.com/en-us/windows/win32/taskschd/monthlytrigger
        Months =    if (($_.Triggers.MonthsOfYear) -ne $null) 
                        {($_.Triggers.MonthsOfYear)}     
                    else 
                        {"N/A"} 
        DaysOfMonth =   if (($_.Triggers.DaysOfMonth) -ne $null)
                            {($_.Triggers.DaysOfMonth)}
                        else 
                            {"N/A"} 
        
        Repetition = ($_.Triggers.Repetition | Select-Object -ExpandProperty interval)
        Duration = ($_.Triggers.Repetition | Select-Object -ExpandProperty duration)     

        }
     } | Export-Csv -Path $outcsv -Encoding UTF8 -NoTypeInformation