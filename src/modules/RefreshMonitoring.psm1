function Get-RefreshInventory {
    if ($null -ne $script:RefreshInventory) {
        return $script:RefreshInventory
    }

    $script:RefreshInventory = @()
    $modelEntries = @(
        foreach ($workspace in @($script:PublicWorkspaces | Where-Object { $_.Type -in @('Workspace', 'Group') })) {
            foreach ($model in @($workspace.Datasets | Where-Object { $null -ne $_ -and $_.Id })) {
                [PSCustomObject]@{ Workspace = $workspace; Model = $model }
            }
        }
    )
    Write-Host ("Collecting refresh history for {0} non-personal semantic models..." -f $modelEntries.Count)
    for ($modelIndex = 0; $modelIndex -lt $modelEntries.Count; $modelIndex++) {
        $entry = $modelEntries[$modelIndex]
        Write-Host ("[History {0}/{1}] {2} / {3}" -f ($modelIndex + 1), $modelEntries.Count, $entry.Workspace.Name, $entry.Model.Name)
        Write-Progress -Activity 'Retrieving refresh history' -Status ("Semantic model {0} of {1}" -f ($modelIndex + 1), $modelEntries.Count) -PercentComplete ([int](100 * ($modelIndex + 1) / [math]::Max(1, $modelEntries.Count)))
        try {
            $url = "groups/$($entry.Workspace.Id)/datasets/$($entry.Model.Id)/refreshes?`$top=60"
            $result = Invoke-PbiApi -Url $url
            $history = @($result.value)
            if ($history.Count -eq 0) {
                $script:RefreshInventory += [PSCustomObject]@{
                    SemanticModel = $entry.Model.Name
                    SemanticModelId = [string]$entry.Model.Id
                    Workspace = $entry.Workspace.Name
                    WorkspaceId = [string]$entry.Workspace.Id
                    Status = 'Never Refreshed'
                    StartTime = $null
                    EndTime = $null
                    Error = $null
                    ServiceExceptionJson = $null
                    RequestId = $null
                    RefreshAttempts = @()
                    RefreshType = $null
                }
            }
            else {
                foreach ($refresh in $history) {
                    $script:RefreshInventory += [PSCustomObject]@{
                        SemanticModel = $entry.Model.Name
                        SemanticModelId = [string]$entry.Model.Id
                        Workspace = $entry.Workspace.Name
                        WorkspaceId = [string]$entry.Workspace.Id
                        Status = [string]$refresh.Status
                        StartTime = $refresh.StartTime
                        EndTime = $refresh.EndTime
                        Error = $refresh.ServiceExceptionJson
                        ServiceExceptionJson = $refresh.ServiceExceptionJson
                        RequestId = $refresh.RequestId
                        RefreshAttempts = @($refresh.RefreshAttempts)
                        RefreshType = $refresh.RefreshType
                    }
                }
            }
        }
        catch {
            $script:RefreshInventory += [PSCustomObject]@{
                SemanticModel = $entry.Model.Name
                SemanticModelId = [string]$entry.Model.Id
                Workspace = $entry.Workspace.Name
                WorkspaceId = [string]$entry.Workspace.Id
                Status = 'Unavailable'
                StartTime = $null
                EndTime = $null
                Error = $_.Exception.Message
                ServiceExceptionJson = $null
                RequestId = $null
                RefreshAttempts = @()
                RefreshType = $null
            }
        }
    }
    Write-Progress -Activity 'Retrieving refresh history' -Completed
    return $script:RefreshInventory
}

function Get-RefreshFailureInfo {
    param ([object]$Refresh)

    $original = [string]$Refresh.ServiceExceptionJson
    $errorCode = $null
    $errorDescription = $null
    if ($original) {
        try {
            $parsed = $original | ConvertFrom-Json -ErrorAction Stop
            $errorCode = [string]$parsed.errorCode
            $errorDescription = [string]$parsed.errorDescription
            if (-not $errorCode) { $errorCode = [string]$parsed.error.code }
            if (-not $errorDescription) { $errorDescription = [string]$parsed.error.message }
        }
        catch { $errorDescription = $original }
    }
    $attemptText = @($Refresh.RefreshAttempts) | ConvertTo-Json -Depth 50 -Compress
    $evidenceText = "$original $($Refresh.Error) $attemptText"
    $category = 'Unknown / Other'
    $reason = 'The available error does not identify a specific cause.'
    $action = 'Review the complete original Power BI error and request ID; validate the next refresh.'

    if ($evidenceText -match '(?i)timeout|timed\s*out|time\s+limit|execution.*expired') {
        $category = 'Timeout'
        $reason = 'Datasource/query operation timed out.'
        $action = 'Review query/source performance and historical refresh duration; validate the next refresh.'
    }
    elseif ($evidenceText -match '(?i)credential|authentication|invalidpassword|login failed|oauth|token.*expired') {
        $category = 'Credentials / Authentication'
        $reason = 'Datasource credentials or authentication were rejected or expired.'
        $action = 'Verify or update the affected datasource credentials, then validate the next refresh.'
    }
    elseif ($evidenceText -match '(?i)throttl|capacity|resource.*limit|out.of.memory|memory.*limit|too many requests|429') {
        $category = 'Capacity / Resources / Throttling'
        $reason = 'The error indicates a capacity, resource or throttling condition.'
        $action = 'Investigate capacity utilization and resource limits; review refresh concurrency.'
    }
    elseif ($evidenceText -match '(?i)permission|authoriz|access.denied|forbidden|403') {
        $category = 'Permissions / Authorization'
        $reason = 'Access to a required resource was denied.'
        $action = 'Verify datasource/database authorization and the identity used for refresh.'
    }
    elseif ($evidenceText -match '(?i)schema|data.type|type.mismatch|cannot.convert|column.*(missing|not.found)|invalid.column') {
        $category = 'Schema / Data Type'
        $reason = 'The error indicates a schema change or data type mismatch.'
        $action = 'Compare source schema and data types with model and Power Query expectations.'
    }
    elseif ($evidenceText -match '(?i)unsupported|not.supported|configuration|not.configured') {
        $category = 'Unsupported Datasource / Configuration'
        $reason = 'The datasource or refresh configuration is unsupported or incomplete.'
        $action = 'Verify connector support, refresh configuration and datasource mapping.'
    }
    elseif ($evidenceText -match '(?i)connection.*(failed|refused)|server.*(unavailable|not.found)|database.*(unavailable|not.found)|network|host.*not.found|could not.*connect') {
        $category = 'Server / Database Connectivity'
        $reason = 'A server/database connectivity error was reported.'
        $action = 'Check connectivity and database availability from the gateway or refresh environment.'
    }
    elseif ($evidenceText -match '(?i)mashup|power.query|expression.error|query.*error|syntax') {
        $category = 'Query / Power Query / Mashup'
        $reason = 'Power Query/Mashup reported an error; inspect the underlying source error.'
        $action = 'Inspect the original Mashup error and affected query, including its underlying cause.'
    }
    elseif ($evidenceText -match '(?i)gateway') {
        $category = 'Gateway'
        $reason = 'A gateway-related error was reported.'
        $action = 'Check gateway availability, configuration and datasource mapping.'
    }

    [PSCustomObject]@{
        ErrorCode = $errorCode
        ErrorDescription = $errorDescription
        FailureCategory = $category
        Reason = $reason
        RecommendedAction = $action
    }
}

function Get-RefreshDurationMinutes {
    param ([object]$Refresh, [datetimeoffset]$Now = [datetimeoffset]::UtcNow)

    if (-not $Refresh.StartTime) { return $null }
    try {
        $start = [datetimeoffset]::Parse([string]$Refresh.StartTime, [cultureinfo]::InvariantCulture)
        if ($Refresh.EndTime) {
            $end = [datetimeoffset]::Parse([string]$Refresh.EndTime, [cultureinfo]::InvariantCulture)
        }
        elseif ($Refresh.Status -in @('Unknown', 'InProgress', 'NotStarted', 'Queued')) { $end = $Now }
        else { return $null }
        if ($end -lt $start) { return $null }
        return [math]::Round(($end - $start).TotalMinutes, 2)
    }
    catch { return $null }
}

function Get-RefreshSequenceClassification {
    param ([object[]]$History, [int]$SequenceLimit = 5)

    $position = 0
    $datedRows = @(
        foreach ($refresh in @($History | Where-Object { $null -ne $_ })) {
            $timestamp = [datetimeoffset]::MinValue
            $validTimestamp = [datetimeoffset]::TryParse([string]$refresh.StartTime, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$timestamp)
            [PSCustomObject]@{
                Refresh = $refresh
                ValidTimestamp = $validTimestamp
                OrderTicks = $(if ($validTimestamp) { $timestamp.UtcTicks } else { [long]::MaxValue })
                Position = $position
            }
            $position++
        }
    )
    $orderedRows = @($datedRows | Sort-Object @{Expression={$_.OrderTicks};Descending=$true}, Position)
    $orderedHistory = @($orderedRows | ForEach-Object { $_.Refresh })
    $evaluated = @($orderedRows | Select-Object -First $SequenceLimit)
    $statuses = @($evaluated | ForEach-Object { [string]$_.Refresh.Status })
    $classification = 'UNKNOWN / NEEDS REVIEW'
    $reason = 'Refresh history is unavailable or has insufficient usable timestamps/statuses to determine the current condition.'
    $consecutiveFailures = 0
    $isRunning = $false
    if ($evaluated.Count -gt 0 -and @($evaluated | Where-Object { -not $_.ValidTimestamp }).Count -eq 0) {
        $latest = $orderedHistory[0]
        $isRunning = $latest.Status -in @('Unknown', 'InProgress', 'NotStarted', 'Queued') -and -not $latest.EndTime
        if ($latest.Status -eq 'Failed') {
            foreach ($status in $statuses) {
                if ($status -ne 'Failed') { break }
                $consecutiveFailures++
            }
            $classification = 'CURRENTLY FAILED / HIGH RISK'
            $reason = "The latest refresh failed; $consecutiveFailures consecutive failure(s) in the evaluated newest-to-oldest sequence."
        }
        elseif ($isRunning) {
            $classification = 'CURRENTLY RUNNING'
            $reason = 'The latest refresh is in progress; older failures do not make it currently failed.'
        }
        elseif ($latest.Status -eq 'Completed' -and @($statuses | Where-Object { $_ -notin @('Completed', 'Failed') }).Count -eq 0) {
            $transitions = 0
            for ($statusIndex = 1; $statusIndex -lt $statuses.Count; $statusIndex++) {
                if ($statuses[$statusIndex] -ne $statuses[$statusIndex - 1]) { $transitions++ }
            }
            if ($statuses -notcontains 'Failed') {
                $classification = 'HEALTHY'
                $reason = 'The latest refresh completed; the evaluated sequence contains only successful refreshes.'
            }
            elseif ($transitions -ge 2) {
                $classification = 'INTERMITTENT / WARNING'
                $reason = 'The latest refresh completed, but the evaluated sequence alternates between success and failure.'
            }
            else {
                $classification = 'RECOVERED'
                $reason = 'Previous refreshes failed, but the latest refresh completed successfully.'
            }
        }
    }
    [PSCustomObject]@{
        Classification = $classification
        Reason = $reason
        OrderedHistory = $orderedHistory
        EvaluatedSequence = $statuses -join ' -> '
        ConsecutiveFailures = $consecutiveFailures
        IsRunning = [bool]$isRunning
    }
}

function Get-RefreshHealthInventory {
    $snapshotTime = [datetimeoffset]::UtcNow
    $recentCutoff = $snapshotTime.AddDays(-7)
    $history = @(Get-RefreshInventory)
    $healthRows = [System.Collections.Generic.List[object]]::new()
    $detailRows = [System.Collections.Generic.List[object]]::new()
    $impactRows = [System.Collections.Generic.List[object]]::new()
    $gatewayLookup = @{}
    $publicWorkspaces = @($script:PublicWorkspaces | Where-Object { $_.Type -in @('Workspace', 'Group') })
    $reports = @((Get-ConsoleContent -WorkspaceSet $publicWorkspaces).Reports)

    $modelGroups = @($history | Group-Object WorkspaceId, SemanticModelId)
    $healthModelIndex = 0
    foreach ($modelGroup in $modelGroups) {
        $healthModelIndex++
        $sequence = Get-RefreshSequenceClassification -History @($modelGroup.Group)
        $modelHistory = @($sequence.OrderedHistory)
        $latest = $modelHistory[0]
        Write-Host ("[Health {0}/{1}] {2} / {3}: checking datasources..." -f $healthModelIndex, $modelGroups.Count, $latest.Workspace, $latest.SemanticModel)
        $modelUrl = "groups/$($latest.WorkspaceId)/datasets/$($latest.SemanticModelId)"
        $gaps = [System.Collections.Generic.List[string]]::new()
        $sources = @()
        $gatewayNames = [System.Collections.Generic.List[string]]::new()
        $knownGatewayIssue = $false
        try {
            $sourceResult = Invoke-PbiApi -Url "$modelUrl/datasources"
            $sources = @($sourceResult.value | Where-Object { $null -ne $_ })
            if ($sources.Count -eq 0) { $gaps.Add('No datasource metadata returned.') }
        }
        catch { $gaps.Add("Datasource metadata unavailable: $($_.Exception.Message)") }

        foreach ($gatewayId in @($sources.gatewayId | Where-Object { $_ -and $_ -ne '00000000-0000-0000-0000-000000000000' } | Sort-Object -Unique)) {
            if (-not $gatewayLookup.ContainsKey([string]$gatewayId)) {
                Write-Host ("  Checking gateway {0}..." -f $gatewayId)
                try {
                    $gatewayLookup[[string]$gatewayId] = [PSCustomObject]@{ Gateway = (Invoke-PbiApi -Url "gateways/$gatewayId"); Error = $null }
                }
                catch { $gatewayLookup[[string]$gatewayId] = [PSCustomObject]@{ Gateway = $null; Error = $_.Exception.Message } }
            }
            $gatewayEntry = $gatewayLookup[[string]$gatewayId]
            if ($gatewayEntry.Error) {
                $gatewayNames.Add([string]$gatewayId)
                $gaps.Add("Gateway $gatewayId metadata unavailable: $($gatewayEntry.Error)")
            }
            else {
                if ($gatewayEntry.Gateway.name) { $gatewayNames.Add([string]$gatewayEntry.Gateway.name) }
                else { $gatewayNames.Add([string]$gatewayId) }
                if ($gatewayEntry.Gateway.gatewayStatus -match '^(Offline|Disconnected|NotReachable)$') { $knownGatewayIssue = $true }
                elseif (-not $gatewayEntry.Gateway.gatewayStatus) { $gaps.Add("Gateway $gatewayId availability state was not returned; availability cannot be validated.") }
            }
        }

        $schedule = $null
        $scheduleHealth = 'UNKNOWN'
        $workspace = $publicWorkspaces | Where-Object { [string]$_.Id -eq $latest.WorkspaceId } | Select-Object -First 1
        $model = $workspace.Datasets | Where-Object { [string]$_.Id -eq $latest.SemanticModelId } | Select-Object -First 1
        if ($model.IsRefreshable -eq $false) { $scheduleHealth = 'NOT APPLICABLE' }
        else {
            Write-Host '  Checking refresh schedule...'
            try {
                $schedule = Invoke-PbiApi -Url "$modelUrl/refreshSchedule"
                if ($null -eq $schedule.enabled) { $gaps.Add('Refresh schedule enabled state was not returned.') }
                elseif (-not $schedule.enabled) { $scheduleHealth = 'DISABLED' }
                elseif (@($schedule.days | Where-Object { $_ }).Count -eq 0 -or @($schedule.times | Where-Object { $_ }).Count -eq 0 -or -not $schedule.localTimeZoneId) {
                    $scheduleHealth = 'WARNING'
                    $gaps.Add('Enabled schedule has incomplete days, times or time zone metadata.')
                }
                else { $scheduleHealth = 'ENABLED' }
            }
            catch { $gaps.Add("Refresh schedule unavailable: $($_.Exception.Message)") }
        }

        $recent = @($modelHistory | Where-Object {
            try { $_.StartTime -and [datetimeoffset]::Parse([string]$_.StartTime, [cultureinfo]::InvariantCulture) -ge $recentCutoff }
            catch { $false }
        } | Select-Object -First 5)
        $failures = @($recent | Where-Object Status -eq 'Failed')
        $lastSuccess = $modelHistory | Where-Object Status -eq 'Completed' | Select-Object -First 1
        $lastFailure = $modelHistory | Where-Object Status -eq 'Failed' | Select-Object -First 1
        $evidence = [System.Collections.Generic.List[string]]::new()
        $actions = [System.Collections.Generic.List[string]]::new()
        $color = 'GREY'
        $health = 'UNKNOWN'
        $risk = 'UNKNOWN'
        $isRunning = $sequence.IsRunning
        $evidence.Add("Newest -> oldest: $($sequence.EvaluatedSequence). $($sequence.Reason)")
        switch ($sequence.Classification) {
            'HEALTHY' { $color = 'GREEN'; $health = 'HEALTHY'; $risk = 'LOW' }
            'RECOVERED' { $color = 'GREEN'; $health = 'RECOVERED'; $risk = 'LOW' }
            'INTERMITTENT / WARNING' { $color = 'YELLOW'; $health = 'INTERMITTENT / WARNING'; $risk = 'WARNING' }
            'CURRENTLY RUNNING' {
                $health = 'CURRENTLY RUNNING'
                if ($lastSuccess) { $color = 'GREEN'; $risk = 'LOW' }
                if ($failures.Count -gt 0) { $color = 'YELLOW'; $risk = 'WARNING' }
            }
            'CURRENTLY FAILED / HIGH RISK' { $color = 'RED'; $health = 'CURRENTLY FAILED / HIGH RISK'; $risk = 'HIGH RISK' }
        }
        if ($isRunning -and -not $lastSuccess) {
            $evidence.Add('A refresh is running, but no prior successful refresh is available to establish health or duration baseline.')
            $actions.Add('Monitor completion and establish a successful-refresh baseline.')
        }
        if ($latest.Status -eq 'Completed') {
            try {
                if (-not $latest.StartTime) { throw 'Missing start timestamp.' }
                if ([datetimeoffset]::Parse([string]$latest.StartTime, [cultureinfo]::InvariantCulture) -lt $recentCutoff) {
                    $color = 'YELLOW'; $health = 'WARNING'; $risk = 'WARNING'
                    $evidence.Add('Latest recorded success is older than seven days; compare this monitoring threshold with the intended schedule cadence.')
                    $actions.Add('Verify the intended refresh frequency and whether more recent refresh history should exist.')
                }
            }
            catch { $gaps.Add('Latest successful refresh has a missing or invalid start timestamp.') }
        }
        if ($failures.Count -gt 0) {
            $evidence.Add("$($failures.Count) historical failure(s) among the latest five refreshes within seven days; current condition is determined by the newest refresh.")
            foreach ($failure in $failures) {
                $diagnosis = Get-RefreshFailureInfo -Refresh $failure
                $evidence.Add("Previous failure at $($failure.StartTime): $($diagnosis.FailureCategory); $($diagnosis.ErrorCode)")
                $actions.Add($diagnosis.RecommendedAction)
            }
        }
        if ($knownGatewayIssue) {
            $color = 'ORANGE'; $health = 'HIGH RISK'; $risk = 'HIGH RISK'
            $evidence.Add('Gateway metadata explicitly reports an offline or disconnected gateway.')
            $actions.Add('Check gateway availability, configuration and datasource mapping.')
        }
        $latestDiagnosis = Get-RefreshFailureInfo -Refresh $latest
        if ($sequence.Classification -eq 'CURRENTLY FAILED / HIGH RISK') {
            $color = 'RED'; $health = 'CURRENTLY FAILED / HIGH RISK'; $risk = 'HIGH RISK'
            $evidence.Add('Power BI reports the latest refresh as Failed; the next outcome is not yet known.')
            $actions.Add($latestDiagnosis.RecommendedAction)
        }
        $duration = Get-RefreshDurationMinutes -Refresh $latest -Now $snapshotTime
        $baselineDurations = @($modelHistory | Where-Object { $_ -ne $latest -and $_.Status -eq 'Completed' } | Select-Object -First 10 | ForEach-Object { Get-RefreshDurationMinutes -Refresh $_ } | Where-Object { $null -ne $_ -and $_ -gt 0 } | Sort-Object)
        $baseline = $null
        if ($baselineDurations.Count -ge 3) {
            $middle = [int][math]::Floor($baselineDurations.Count / 2)
            if ($baselineDurations.Count % 2) { $baseline = $baselineDurations[$middle] }
            else { $baseline = ($baselineDurations[$middle - 1] + $baselineDurations[$middle]) / 2 }
            if ($null -ne $duration -and $duration -gt [math]::Max(5, 2 * $baseline)) {
                if ($risk -ne 'HIGH RISK') { $color = 'YELLOW'; $health = 'WARNING'; $risk = 'WARNING' }
                $evidence.Add("Duration $duration minutes exceeds twice the median ($baseline minutes) of $($baselineDurations.Count) prior successful refreshes, with a five-minute minimum threshold.")
                $actions.Add('Review query/source performance and historical refresh duration.')
            }
        }
        $attempts = @($latest.RefreshAttempts | Where-Object { $null -ne $_ })
        $repeatedAttemptTypes = @($attempts | Where-Object { $_.type } | Group-Object type | Where-Object Count -gt 1)
        $attemptErrors = @($attempts | Where-Object { $_.serviceExceptionJson })
        if ($repeatedAttemptTypes.Count -gt 0 -or $attemptErrors.Count -gt 0) {
            if ($risk -notin @('HIGH RISK', 'WARNING')) { $color = 'YELLOW'; $health = 'WARNING'; $risk = 'WARNING' }
            $repeatedTypes = @($repeatedAttemptTypes | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join '; '
            $evidence.Add("Latest refresh contains repeated attempt types [$repeatedTypes] or attempt-level errors ($($attemptErrors.Count)); inspect the original records.")
            $actions.Add('Inspect the original refresh attempt records for repeated operations and underlying errors.')
        }
        if ($sequence.Classification -eq 'UNKNOWN / NEEDS REVIEW') {
            if ($risk -notin @('HIGH RISK', 'WARNING')) { $color = 'GREY'; $health = 'UNKNOWN'; $risk = 'UNKNOWN' }
            $evidence.Add("Refresh outcome cannot be assessed: $($latest.Status). $($latest.Error)")
            $actions.Add('Verify refresh applicability and monitoring access, then retrieve fresh history.')
        }
        if ($gaps.Count -gt 0) {
            if ($risk -eq 'LOW') { $color = 'GREY'; $health = 'UNKNOWN'; $risk = 'UNKNOWN' }
            foreach ($gap in $gaps) { $evidence.Add($gap) }
            $actions.Add('Resolve metadata/monitoring gaps; unavailable metadata does not prove a datasource failure.')
        }
        if ($evidence.Count -eq 0) { $evidence.Add('No known failure condition in the retrieved refresh history; this is not a guarantee of future success.') }
        if ($actions.Count -eq 0) { $actions.Add('Monitor the next refresh and compare its outcome and duration with this snapshot.') }
        $scheduleAssessment = $scheduleHealth
        if ($scheduleHealth -eq 'ENABLED' -and $risk -ne 'LOW') { $scheduleAssessment = "$scheduleHealth / $risk" }
        $connections = @(
            foreach ($source in $sources) {
                $connection = $source.connectionDetails
                if ($connection -is [string]) {
                    try { $connection = $connection | ConvertFrom-Json -ErrorAction Stop }
                    catch { $connection = $null }
                }
                [PSCustomObject]@{ Server = $connection.server; Database = $connection.database; DatasourceType = $source.datasourceType; DatasourceId = $source.datasourceId; GatewayId = $source.gatewayId }
            }
        )
        $server = @($connections.Server | Where-Object { $_ } | Sort-Object -Unique) -join '; '
        $database = @($connections.Database | Where-Object { $_ } | Sort-Object -Unique) -join '; '
        $datasourceType = @($connections.DatasourceType | Where-Object { $_ } | Sort-Object -Unique) -join '; '
        $gateway = @($gatewayNames | Sort-Object -Unique) -join '; '
        $affectedReports = @($reports | Where-Object { $_.WorkspaceId -eq $latest.WorkspaceId -and $_.SemanticModel -eq $latest.SemanticModelId })
        $healthRow = [PSCustomObject]@{
            Workspace = $latest.Workspace; WorkspaceId = $latest.WorkspaceId
            SemanticModel = $latest.SemanticModel; SemanticModelId = $latest.SemanticModelId
            HealthColor = $color; Health = $health; CurrentRiskLevel = $risk
            CurrentClassification = $sequence.Classification
            ClassificationReason = $sequence.Reason
            EvaluatedSequence = $sequence.EvaluatedSequence
            ConsecutiveFailures = $sequence.ConsecutiveFailures
            Evidence = @($evidence | Sort-Object -Unique) -join ' | '
            LastSuccessfulRefresh = $lastSuccess.StartTime; LastFailedRefresh = $lastFailure.StartTime
            FailureCount = $failures.Count; LatestStatus = $latest.Status; IsRunning = [bool]$isRunning
            StartTime = $latest.StartTime; EndTime = $latest.EndTime; DurationMinutes = $duration
            BaselineDurationMinutes = $baseline; RefreshType = $latest.RefreshType
            Server = $server; Database = $database; DatasourceType = $datasourceType; Gateway = $gateway
            DatasourceMappings = ConvertTo-Json -InputObject @($connections) -Depth 10 -Compress
            ScheduleHealth = $scheduleAssessment; ScheduleEnabled = $schedule.enabled
            ScheduleDays = @($schedule.days) -join '; '; ScheduleTimes = @($schedule.times) -join '; '
            ScheduleTimeZone = $schedule.localTimeZoneId; ScheduleNotifyOption = $schedule.notifyOption
            MonitoringGaps = $gaps -join ' | '
            ErrorCode = $latestDiagnosis.ErrorCode; ErrorDescription = $latestDiagnosis.ErrorDescription
            FailureCategory = $(if ($latest.Status -eq 'Failed') { $latestDiagnosis.FailureCategory } else { $null })
            Reason = $(if ($latest.Status -eq 'Failed') { $latestDiagnosis.Reason } else { $null })
            ServiceExceptionJson = $latest.ServiceExceptionJson; RequestId = $latest.RequestId
            RefreshAttempts = ConvertTo-Json -InputObject @($attempts) -Depth 50 -Compress
            RecentFailureDetails = ConvertTo-Json -InputObject @($failures | Select-Object StartTime, EndTime, ServiceExceptionJson, RequestId, RefreshAttempts) -Depth 50 -Compress
            FailureHistoryDetails = ConvertTo-Json -InputObject @($modelHistory | Where-Object Status -eq 'Failed' | Select-Object StartTime, EndTime, ServiceExceptionJson, RequestId, RefreshAttempts) -Depth 50 -Compress
            CredentialAssessment = 'No live credential test is performed; credential findings are based on returned Power BI errors.'
            AffectedReports = @($affectedReports.Name) -join '; '
            RecommendedAction = @($actions | Sort-Object -Unique) -join ' | '
            Validation = 'After an administrator acts, reopen Refresh Monitoring to retrieve a new snapshot; confirm a subsequent Completed refresh and compare duration. No automatic changes are made.'
            SnapshotTimeUtc = $snapshotTime.ToString('o')
        }
        $healthRows.Add($healthRow)
        foreach ($refresh in $modelHistory) {
            $diagnosis = Get-RefreshFailureInfo -Refresh $refresh
            $detailRows.Add([PSCustomObject]@{
                Workspace = $refresh.Workspace; WorkspaceId = $refresh.WorkspaceId
                SemanticModel = $refresh.SemanticModel; SemanticModelId = $refresh.SemanticModelId
                Server = $server; Database = $database; DatasourceType = $datasourceType; Gateway = $gateway
                DatasourceMappings = $healthRow.DatasourceMappings
                Status = $refresh.Status; RefreshType = $refresh.RefreshType
                StartTime = $refresh.StartTime; EndTime = $refresh.EndTime
                DurationMinutes = Get-RefreshDurationMinutes -Refresh $refresh -Now $snapshotTime
                ErrorCode = $diagnosis.ErrorCode; ErrorDescription = $diagnosis.ErrorDescription
                FailureCategory = $(if ($refresh.Status -eq 'Failed') { $diagnosis.FailureCategory } else { $null })
                Reason = $(if ($refresh.Status -eq 'Failed') { $diagnosis.Reason } else { $refresh.Error })
                RecommendedAction = $(if ($refresh.Status -eq 'Failed') { $diagnosis.RecommendedAction } else { $null })
                RequestId = $refresh.RequestId; ServiceExceptionJson = $refresh.ServiceExceptionJson
                RefreshAttempts = ConvertTo-Json -InputObject @($refresh.RefreshAttempts) -Depth 50 -Compress
                AffectedReports = $healthRow.AffectedReports
            })
        }
        if ($risk -ne 'LOW') {
            foreach ($report in $affectedReports) {
                $impactRows.Add([PSCustomObject]@{ Server = $server; Database = $database; DatasourceType = $datasourceType; DatasourceMappings = $healthRow.DatasourceMappings; Gateway = $gateway; SemanticModel = $latest.SemanticModel; SemanticModelId = $latest.SemanticModelId; Workspace = $latest.Workspace; WorkspaceId = $latest.WorkspaceId; Report = $report.Name; ReportId = $report.Id; CurrentRiskLevel = $risk; Evidence = $healthRow.Evidence; RecommendedAction = $healthRow.RecommendedAction })
            }
        }
    }
    [PSCustomObject]@{ Health = $healthRows.ToArray(); History = $detailRows.ToArray(); Impact = $impactRows.ToArray(); SnapshotTimeUtc = $snapshotTime.ToString('o') }
}

function Write-RefreshDetails {
    param ([string]$Title, [object[]]$Items, [string[]]$Properties = @('*'))

    Write-Host "`n$Title"
    if (@($Items).Count -eq 0) { Write-Host 'None found in the retrieved public-workspace snapshot.' }
    else { $Items | Format-List -Property $Properties | Out-Host }
}

function Show-RefreshSection {
    $options = @(
        'Refresh Health Summary', 'Currently Failed', 'At Risk Before Next Refresh',
        'Currently Running Refreshes', 'Scheduled Refresh Health', 'Failure Reasons',
        'Credential / Authentication Issues', 'Gateway Issues', 'Timeout / Performance Issues',
        'Server / Database Issues', 'Query / Power Query Issues', 'Capacity / Resource Issues',
        'Affected Reports', 'Recommended Actions', 'Refresh History',
        'Semantic Models Never Refreshed', 'Search Semantic Model', 'Export Refresh Health',
        'Recovered', 'Intermittent / Warning', 'Healthy', 'Unknown / Needs Review',
        'Full Failure History', 'High Risk', 'Back'
    )
    $choice = Read-ConsoleMenu -Title 'REFRESH MONITORING - NON-PERSONAL WORKSPACES' -Options $options
    if ($choice -eq 25) { return }
    Write-Host ("`nPreparing {0}. Results will appear in this terminal after collection completes." -f $options[$choice - 1])
    $script:RefreshInventory = $null
    $snapshot = Get-RefreshHealthInventory
    $health = @($snapshot.Health)
    $history = @($snapshot.History)
    $failed = @($history | Where-Object Status -eq 'Failed')
    $currentlyFailed = @($health | Where-Object CurrentClassification -eq 'CURRENTLY FAILED / HIGH RISK')
    $issues = @($health | Where-Object CurrentRiskLevel -ne 'LOW')
    Write-Host ("Collection complete: {0} models assessed; {1} with warnings, high risk or unknown information." -f $health.Count, $issues.Count)
    Write-Host ("Snapshot UTC: {0}. Current classification: newest -> oldest, latest five records. Historical evidence window: latest five refreshes within seven days. Duration baseline: up to ten prior successful refreshes, minimum three." -f $snapshot.SnapshotTimeUtc)
    Write-Host 'Read-only assessment. WARNING/HIGH RISK are evidence-based indicators, not predictions of certain failure. Reopen this section after an action to validate a new snapshot.'
    switch ($choice) {
        1 {
            foreach ($level in @('GREEN', 'YELLOW', 'ORANGE', 'RED', 'GREY')) {
                Write-Host ("{0,-8} {1}" -f $level, @($health | Where-Object HealthColor -eq $level).Count)
            }
            foreach ($classification in @('CURRENTLY FAILED / HIGH RISK', 'RECOVERED', 'INTERMITTENT / WARNING', 'HEALTHY', 'CURRENTLY RUNNING', 'UNKNOWN / NEEDS REVIEW')) {
                Write-Host ("{0,-30} {1}" -f $classification, @($health | Where-Object CurrentClassification -eq $classification).Count)
            }
            Write-DetailTable -Title 'REFRESH HEALTH SUMMARY' -Items $health -Properties @('Workspace', 'SemanticModel', 'CurrentClassification', 'HealthColor', 'CurrentRiskLevel', 'LatestStatus')
        }
        2 { Write-RefreshDetails -Title 'CURRENTLY FAILED - LATEST REFRESH ONLY' -Items $currentlyFailed }
        3 { Write-RefreshDetails -Title 'AT RISK BEFORE NEXT REFRESH' -Items $issues -Properties @('SemanticModel', 'Workspace', 'CurrentClassification', 'LatestStatus', 'EvaluatedSequence', 'ConsecutiveFailures', 'HealthColor', 'CurrentRiskLevel', 'Evidence', 'LastSuccessfulRefresh', 'LastFailedRefresh', 'FailureCount', 'ScheduleHealth', 'RecommendedAction', 'Validation') }
        4 { Write-RefreshDetails -Title 'CURRENTLY RUNNING REFRESHES' -Items @($health | Where-Object IsRunning -eq $true) -Properties @('Workspace', 'SemanticModel', 'CurrentClassification', 'LatestStatus', 'StartTime', 'DurationMinutes', 'BaselineDurationMinutes', 'CurrentRiskLevel', 'Evidence', 'RecommendedAction') }
        5 { Write-RefreshDetails -Title 'SCHEDULED REFRESH HEALTH' -Items $health -Properties @('Workspace', 'SemanticModel', 'ScheduleHealth', 'ScheduleEnabled', 'ScheduleDays', 'ScheduleTimes', 'ScheduleTimeZone', 'ScheduleNotifyOption', 'CurrentRiskLevel', 'Evidence', 'RecommendedAction') }
        6 {
            Write-DetailTable -Title 'FAILURE REASON COUNTS' -Items @($failed | Group-Object FailureCategory | Select-Object @{Name='FailureCategory';Expression={$_.Name}}, Count)
            Write-RefreshDetails -Title 'FAILURE REASONS AND ORIGINAL ERRORS' -Items $failed
        }
        { $_ -in @(7, 8, 9, 10, 11, 12) } {
            $categoryMap = @{
                7 = @('Credentials / Authentication', 'Permissions / Authorization')
                8 = @('Gateway')
                9 = @('Timeout')
                10 = @('Server / Database Connectivity')
                11 = @('Query / Power Query / Mashup', 'Schema / Data Type', 'Unsupported Datasource / Configuration')
                12 = @('Capacity / Resources / Throttling')
            }
            Write-RefreshDetails -Title $options[$choice - 1].ToUpperInvariant() -Items @($failed | Where-Object { $_.FailureCategory -in $categoryMap[[int]$choice] })
            if ($choice -eq 8) {
                Write-RefreshDetails -Title 'CURRENT GATEWAY WARNINGS / MONITORING GAPS' -Items @($issues | Where-Object { $_.Evidence -match '(?i)gateway' })
            }
            if ($choice -eq 9) {
                Write-RefreshDetails -Title 'ABNORMAL DURATION WARNINGS' -Items @($issues | Where-Object { $_.Evidence -match 'exceeds twice the median' })
            }
        }
        13 {
            Write-RefreshDetails -Title 'AFFECTED REPORTS' -Items @($snapshot.Impact)
            Write-RefreshDetails -Title 'ISSUES WITH NO LINKED REPORT METADATA' -Items @($issues | Where-Object { -not $_.AffectedReports }) -Properties @('Workspace', 'SemanticModel', 'Server', 'Database', 'DatasourceMappings', 'CurrentRiskLevel', 'Evidence', 'RecommendedAction')
        }
        14 { Write-RefreshDetails -Title 'RECOMMENDED ACTIONS - ADMINISTRATOR REVIEW REQUIRED' -Items $issues -Properties @('Workspace', 'SemanticModel', 'Server', 'Database', 'Gateway', 'CurrentRiskLevel', 'Evidence', 'AffectedReports', 'RecommendedAction', 'Validation') }
        15 { Write-RefreshDetails -Title 'REFRESH HISTORY (UP TO 60 PER MODEL)' -Items $history }
        16 { Write-RefreshDetails -Title 'NO REFRESH HISTORY RETURNED - VERIFY APPLICABILITY AND RETENTION' -Items @($health | Where-Object LatestStatus -eq 'Never Refreshed') }
        17 { $term = Read-Host 'Semantic model name or ID'; Write-RefreshDetails -Title 'REFRESH HEALTH SEARCH' -Items @($health | Where-Object { $_.SemanticModel -like "*$term*" -or $_.SemanticModelId -eq $term }) }
        18 { Export-ConsoleRows -Rows $health -BaseName 'PowerBI_RefreshHealth' }
        19 { Write-RefreshDetails -Title 'RECOVERED - LATEST REFRESH COMPLETED' -Items @($health | Where-Object CurrentClassification -eq 'RECOVERED') }
        20 { Write-RefreshDetails -Title 'INTERMITTENT / WARNING' -Items @($health | Where-Object { $_.CurrentClassification -eq 'INTERMITTENT / WARNING' -or $_.CurrentRiskLevel -eq 'WARNING' }) }
        21 { Write-RefreshDetails -Title 'HEALTHY' -Items @($health | Where-Object { $_.CurrentClassification -eq 'HEALTHY' -and $_.CurrentRiskLevel -eq 'LOW' }) }
        22 { Write-RefreshDetails -Title 'UNKNOWN / NEEDS REVIEW' -Items @($health | Where-Object { $_.CurrentClassification -eq 'UNKNOWN / NEEDS REVIEW' -or $_.CurrentRiskLevel -eq 'UNKNOWN' }) }
        23 { Write-RefreshDetails -Title 'FULL FAILURE HISTORY - ALL RETRIEVED FAILURES, INCLUDING RECOVERED MODELS' -Items $failed }
        24 { Write-RefreshDetails -Title 'HIGH RISK - EVIDENCE-BASED ASSESSMENT' -Items @($health | Where-Object CurrentRiskLevel -eq 'HIGH RISK') }
    }
}
