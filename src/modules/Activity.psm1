function Get-ActivityInventory {
    param ([datetime]$Date = [datetime]::UtcNow)

    $utcDate = $Date.ToUniversalTime()
    $start = [uri]::EscapeDataString("'$(Get-Date -Date $utcDate.Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ")'")
    $end = [uri]::EscapeDataString("'$(Get-Date -Date $utcDate -Format "yyyy-MM-ddTHH:mm:ss.fffZ")'")
    $url = "admin/activityevents?startDateTime=$start&endDateTime=$end"
    $events = [System.Collections.Generic.List[object]]::new()
    $pageCount = 0
    do {
        $result = Invoke-PbiApi -Url $url
        foreach ($event in @($result.activityEventEntities)) {
            $events.Add($event)
        }
        $token = [string]$result.continuationToken
        $pageCount++
        if ($token -and $pageCount -lt 100) {
            $url = 'admin/activityevents?continuationToken=' + [uri]::EscapeDataString($token)
        }
        else {
            $url = $null
        }
    } while ($url)

    return $events.ToArray()
}

function Show-ActivitySection {
    $choice = Read-ConsoleMenu -Title 'ACTIVITY MONITORING' -Options @('Activity Summary', 'Report Views', 'User Activity', 'Workspace Activity', 'Semantic Model Activity', 'Search User', 'Search Report', 'Export Activity', 'Back')
    try {
        $dateText = Read-Host 'UTC date (yyyy-MM-dd; Enter for today)'
        $date = if ([string]::IsNullOrWhiteSpace($dateText)) { [datetime]::UtcNow } else { [datetime]::ParseExact($dateText, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture).ToUniversalTime() }
        $events = @(Get-ActivityInventory -Date $date)
    }
    catch {
        Write-Error "Activity events unavailable. The admin activity API requires Fabric administrator access and a date in the last 28 days: $_"
        return
    }
    switch ($choice) {
        1 { Write-Host ("Activity events retrieved: {0}" -f $events.Count); Write-DetailTable -Title 'ACTIVITY BY OPERATION' -Items @($events | Group-Object Activity | Select-Object @{Name='Activity';Expression={$_.Name}}, Count) }
        2 { Write-DetailTable -Title 'REPORT VIEWS' -Items @($events | Where-Object Activity -match 'ViewReport') }
        3 { Write-DetailTable -Title 'USER ACTIVITY' -Items $events }
        4 { Write-DetailTable -Title 'WORKSPACE ACTIVITY' -Items @($events | Select-Object CreationTime, UserId, Activity, WorkSpaceName, WorkspaceId, ItemName) }
        5 { Write-DetailTable -Title 'SEMANTIC MODEL ACTIVITY' -Items @($events | Select-Object CreationTime, UserId, Activity, WorkSpaceName, DatasetName, DatasetId) }
        6 { $term = Read-Host 'User'; Write-DetailTable -Title 'USER ACTIVITY SEARCH' -Items @($events | Where-Object UserId -like "*$term*") }
        7 { $term = Read-Host 'Report'; Write-DetailTable -Title 'REPORT ACTIVITY SEARCH' -Items @($events | Where-Object { $_.ReportName -like "*$term*" -or $_.ItemName -like "*$term*" }) }
        8 { Export-ConsoleRows -Rows $events -BaseName 'PowerBI_Activity' }
    }
}
