# Import Settings PowerShell script
. ($PSScriptRoot + "\Settings.ps1")

# Initialize global variables
$currentAvailability = $null
$currentActivity = $null
$previousAvailability = $null
$previousActivity = $null
$currentCalendarStatus = $null
$previousCalendarStatus = $null
$currentNextMeeting = $null
$previousNextMeeting = $null
$ActivityIcon = $iconNotInACall

function Get-LatestLogFile {
    try {
        $logFiles = Get-ChildItem -Path $logDirPath -Filter "MSTeams_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 0) {
            return $logFiles[0].FullName
        } else {
            Write-Error "No log files found."
            return $null
        }
    } catch {
        Write-Error "Failed to fetch log files: $_"
        return $null
    }
}

Function Find-LatestAvailability {
    Param ([string]$logFilePath)

    $logFileContent = Get-Content -Path $logFilePath -ReadCount 0 | Select-Object -Last 1000
    [array]::Reverse($logFileContent)
    foreach ($line in $logFileContent) {
        $timestampMatches = [regex]::Matches($line, $timestampPattern)
        $availabilityMatches = [regex]::Matches($line, $availabilityPattern)

        if ($timestampMatches.Count -gt 0 -and $availabilityMatches.Count -gt 0) {
            $availabilitytimestamp = $timestampMatches[0].Value
            $script:currentAvailability = $availabilityMatches[0].Groups[1].Value
            Write-Host "Found availability: $script:currentAvailability at $availabilitytimestamp"
            return "Timestamp: $availabilitytimestamp, Availability: $script:currentAvailability"
        }
    }
    Write-Host "No availability found in the recent log entries."
    return $null
}

Function Find-LatestActivity {
    Param ([string]$logFilePath)

    $logFileContent = Get-Content -Path $logFilePath -ReadCount 0 | Select-Object -Last 1000
    [array]::Reverse($logFileContent)
    foreach ($line in $logFileContent) {
        $timestampMatches = [regex]::Matches($line, $timestampPattern)
        $activityMatches = [regex]::Matches($line, $activityPattern)

        if ($timestampMatches.Count -gt 0 -and $activityMatches.Count -gt 0) {
            $activitytimestamp = $timestampMatches[0].Value
            $script:currentActivity = $activityMatches[0].Groups[1].Value
            Write-Host "Found activity: $script:currentActivity at $activitytimestamp"
            return "Timestamp: $activitytimestamp, Activity: $script:currentActivity"
        }
    }
    Write-Host "No activity found in the recent log entries."
    return $null
}

Function Set-ActivityIcon {
    if ($script:currentActivity -eq "VeryActive") {
        $script:ActivityIcon = $iconInACall
    } else {
        $script:ActivityIcon = $iconNotInACall
    }
}

Function Get-CalendarStatus {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendarFolder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)

    $currentTime = Get-Date
    $items = $calendarFolder.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")

    $currentMeeting = $null
    $nextMeeting = $null

    foreach ($item in $items) {
        if ($item.Start -le $currentTime -and $item.End -gt $currentTime) {
            if ($item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy -or
                $item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olTentative) {
                $currentMeeting = $item
                break
            }
        } elseif ($item.Start -gt $currentTime) {
            if ($item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy -or
                $item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olTentative) {
                $nextMeeting = $item
                break
            }
        }
    }

    if ($currentMeeting) {
        $script:currentCalendarStatus = "Busy"
    } else {
        $script:currentCalendarStatus = "Free"
    }

    if ($nextMeeting) {
        $script:currentNextMeeting = $nextMeeting.Start
    } else {
        $script:currentNextMeeting = "No upcoming meetings today"
    }

    Write-Host "Calendar status: $script:currentCalendarStatus"
    Write-Host "Next meeting: $script:currentNextMeeting"
}

Function Send-LatestActivity {
    if ($script:currentAvailability -ne $script:previousAvailability) {
        $params = @{
            "state" = "$script:currentAvailability";
            "attributes" = @{
                "friendly_name" = "$entityStatusName";
                "icon" = "mdi:microsoft-teams";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated availability to Home Assistant."
        } catch {
            Write-Error "Failed to update availability to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in availability, skipping update."
    }

    if ($script:currentActivity -ne $script:previousActivity) {
        $params = @{
            "state" = "$script:currentActivity";
            "attributes" = @{
                "friendly_name" = "$entityActivityName";
                "icon" = "$script:ActivityIcon";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityActivity" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated activity to Home Assistant."
        } catch {
            Write-Error "Failed to update activity to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in activity, skipping update."
    }

    if ($script:currentCalendarStatus -ne $script:previousCalendarStatus) {
        $params = @{
            "state" = "$script:currentCalendarStatus";
            "attributes" = @{
                "friendly_name" = "$entityCalendarStatusName";
                "icon" = "mdi:calendar";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityCalendarStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated calendar status to Home Assistant."
        } catch {
            Write-Error "Failed to update calendar status to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in calendar status, skipping update."
    }

    if ($script:currentNextMeeting -ne $script:previousNextMeeting) {
        $params = @{
            "state" = "$script:currentNextMeeting";
            "attributes" = @{
                "friendly_name" = "$entityNextMeetingName";
                "icon" = "mdi:calendar-clock";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityNextMeeting" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated next meeting to Home Assistant."
        } catch {
            Write-Error "Failed to update next meeting to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in next meeting, skipping update."
    }
}

Function Monitor-Teams {
    while ($Enable) {
        $latestLogFile = Get-LatestLogFile
        if ($latestLogFile) {
            Find-LatestAvailability -logFilePath $latestLogFile
            Find-LatestActivity -logFilePath $latestLogFile
            Set-ActivityIcon
            Get-CalendarStatus

            Send-LatestActivity

            $previousAvailability = $script:currentAvailability
            $previousActivity = $script:currentActivity
            $previousCalendarStatus = $script:currentCalendarStatus
            $previousNextMeeting = $script:currentNextMeeting
        } else {
            Write-Host "No log file found. Skipping this iteration."
        }

        Start-Sleep -Seconds $refreshDelay
    }
}

# Start monitoring Teams activity and calendar status
Monitor-Teams
