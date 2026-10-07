# Requires MicrosoftPowerBIMgmt.Profile and MicrosoftPowerBIMgmt.Workspaces
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile) -or
    -not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Workspaces)) {
    throw "Power BI module missing. Install it with: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
}

Import-Module MicrosoftPowerBIMgmt.Profile -WarningAction SilentlyContinue
Import-Module MicrosoftPowerBIMgmt.Workspaces -WarningAction SilentlyContinue

try {
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue

    Write-Host "Loading Power BI workspace inventory..."
    $workspaces = @(Get-PowerBIWorkspace -Scope Organization -All -Include All -WarningAction SilentlyContinue)
    $deletedWorkspaces = @(Get-PowerBIWorkspace -Scope Organization -Deleted -All -WarningAction SilentlyContinue)
    $orphanedWorkspaces = @(Get-PowerBIWorkspace -Scope Organization -Orphaned -All -WarningAction SilentlyContinue)

    $deletedOrOrphaned = @($deletedWorkspaces + $orphanedWorkspaces | Sort-Object Id -Unique)
    $publicWorkspaces = @($workspaces | Where-Object { $_.Type -in @('Workspace', 'Group') })
    $personalWorkspaces = @($workspaces | Where-Object { $_.Type -eq 'PersonalGroup' })
    $deletedOrOrphanedPublic = @($deletedOrOrphaned | Where-Object { $_.Type -in @('Workspace', 'Group') })
    $deletedOrOrphanedPersonal = @($deletedOrOrphaned | Where-Object { $_.Type -eq 'PersonalGroup' })
    $selectedWorkspaces = @($workspaces)
    $script:Workspaces = $workspaces
    $script:PublicWorkspaces = $publicWorkspaces
    $script:PersonalWorkspaces = $personalWorkspaces
    $script:DeletedWorkspaces = $deletedWorkspaces
    $script:OrphanedWorkspaces = $orphanedWorkspaces
    $script:Capacities = @()
    $script:ScanWorkspaces = $null
    $script:RefreshInventory = $null
    $script:GatewayInventory = $null
    $script:ActivityInventory = $null
    $script:ServerInventory = $null
    $script:ServerSummaryRows = @()
    $script:OutputFolder = $PSScriptRoot
    $script:LoggingEnabled = $false

    function Write-WorkspaceSummary {
        param (
            [string]$Title,
            [object[]]$CurrentWorkspaces,
            [object[]]$DeletedOrOrphanedWorkspaces
        )

        $knownWorkspaces = @($CurrentWorkspaces + $DeletedOrOrphanedWorkspaces | Sort-Object Id -Unique)
        $activeCount = @($CurrentWorkspaces | Where-Object { $_.State -eq 'Active' }).Count
        $reportCount = 0
        $semanticModelCount = 0
        $dashboardCount = 0
        $dataflowCount = 0

        foreach ($workspace in $CurrentWorkspaces) {
            $reportCount += @($workspace.Reports).Count
            $semanticModelCount += @($workspace.Datasets).Count
            $dashboardCount += @($workspace.Dashboards).Count
            $dataflowCount += @($workspace.Dataflows).Count
        }

        Write-Host "`n$Title"
        Write-Host ("Total                  {0}" -f $knownWorkspaces.Count)
        Write-Host ("Active                 {0}" -f $activeCount)
        Write-Host ("Deleted / orphaned     {0}" -f $DeletedOrOrphanedWorkspaces.Count)
        Write-Host "CONTENT"
        Write-Host ("Reports                {0}" -f $reportCount)
        Write-Host ("Semantic Models        {0}" -f $semanticModelCount)
        Write-Host ("Dashboards             {0}" -f $dashboardCount)
        Write-Host ("Dataflows              {0}" -f $dataflowCount)
    }

    Write-Host "`nPOWER BI SERVICE ADMIN SUMMARY"
    Write-Host "--------------------------------"
    if ($workspaceSelection -eq '1') {
        Write-WorkspaceSummary -Title 'PUBLIC WORKSPACES' -CurrentWorkspaces $publicWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPublic
        $selectedWorkspaces = @($publicWorkspaces)
    }
    elseif ($workspaceSelection -eq '2') {
        Write-WorkspaceSummary -Title 'PERSONAL WORKSPACES' -CurrentWorkspaces $personalWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPersonal
        $selectedWorkspaces = @($personalWorkspaces)
    }
    else {
        Write-WorkspaceSummary -Title 'PUBLIC WORKSPACES' -CurrentWorkspaces $publicWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPublic
        Write-WorkspaceSummary -Title 'PERSONAL WORKSPACES' -CurrentWorkspaces $personalWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPersonal
        $selectedWorkspaces = @($publicWorkspaces + $personalWorkspaces | Sort-Object Id -Unique)
    }

    function Get-ReportKind {
        param ([object]$Report)

        $reportType = [string]$Report.ReportType
        if ([string]::IsNullOrWhiteSpace($reportType)) {
            $reportType = [string]$Report.Type
        }

        if ($reportType -match 'Paginated') {
            return 'PaginatedReport'
        }
        if ($reportType -match 'PowerBI') {
            return 'PowerBIReport'
        }
        if ($Report.Format -eq 'RDL') {
            return 'PaginatedReport'
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$Report.DatasetId)) {
            return 'PowerBIReport'
        }

        return 'PaginatedReport'
    }

    function Get-ArtifactName {
        param ([object]$Item)

        if (-not [string]::IsNullOrWhiteSpace([string]$Item.Name)) {
            return $Item.Name
        }

        return $Item.DisplayName
    }

    function Write-DetailTable {
        param (
            [string]$Title,
            [object[]]$Items,
            [string[]]$Properties
        )

        Write-Host "`n$Title"
        if (@($Items).Count -eq 0) {
            Write-Host "None found."
        }
        elseif ($Properties -and $Properties.Count -gt 0) {
            $Items | Format-Table -Property $Properties -Wrap -AutoSize | Out-String -Width 240 -Stream
        }
        else {
            $Items | Format-Table -Wrap -AutoSize
        }
    }

    function Get-ServerCategory {
        param ([string]$DatasourceType)

        if ($DatasourceType -match '^(?i:Sql|SqlServer)$') { return 'SQL Server' }
        if ($DatasourceType -match '(?i)Analysis.?Services') { return 'Analysis Services' }
        if ($DatasourceType -match '(?i)Oracle') { return 'Oracle' }
        return 'Other'
    }

    function Get-DataSourceScanResults {
        param ([object[]]$Workspaces)

        $scanResults = [System.Collections.Generic.List[object]]::new()
        $failedBatches = 0
        $firstError = $null
        $workspaceIds = @($Workspaces | ForEach-Object { [string]$_.Id } | Where-Object { $_ })
        $batchCount = [math]::Ceiling($workspaceIds.Count / 100.0)

        for ($offset = 0; $offset -lt $workspaceIds.Count; $offset += 100) {
            $batchNumber = [int]($offset / 100) + 1
            $batch = @($workspaceIds | Select-Object -Skip $offset -First 100)
            Write-Progress -Activity 'Building datasource/server inventory' -Status ("Workspace batch {0} of {1}" -f $batchNumber, $batchCount) -PercentComplete ([int](100 * $batchNumber / [math]::Max(1, $batchCount)))

            try {
                $body = @{ workspaces = $batch } | ConvertTo-Json -Depth 5
                $scan = Invoke-PbiApi -Url 'admin/workspaces/getInfo?lineage=true&datasourceDetails=true' -Method Post -Body $body
                $scanId = [string]$scan.id
                if (-not $scanId) {
                    throw 'Metadata scan request returned no scan ID.'
                }

                $status = 'NotStarted'
                for ($attempt = 0; $attempt -lt 60 -and $status -notin @('Succeeded', 'Failed'); $attempt++) {
                    Start-Sleep -Seconds 2
                    $scanStatus = Invoke-PbiApi -Url ("admin/workspaces/scanStatus/{0}" -f $scanId)
                    $status = [string]$scanStatus.status
                }
                if ($status -ne 'Succeeded') {
                    throw ("Metadata scan {0} ended with status {1}." -f $scanId, $status)
                }

                $scanResults.Add((Invoke-PbiApi -Url ("admin/workspaces/scanResult/{0}" -f $scanId)))
            }
            catch {
                $failedBatches++
                if (-not $firstError) {
                    $firstError = $_.Exception.Message
                }
            }
        }
        Write-Progress -Activity 'Building datasource/server inventory' -Completed

        return [PSCustomObject]@{
            Scans = $scanResults.ToArray()
            FailedBatches = $failedBatches
            FirstError = $firstError
        }
    }

    function Get-ServerInventory {
        param ([object[]]$Workspaces)

        $inventory = [System.Collections.Generic.List[object]]::new()
        $scanResult = Get-DataSourceScanResults -Workspaces $Workspaces
        $workspaceById = @{}
        foreach ($workspace in $Workspaces) {
            $workspaceById[[string]$workspace.Id] = $workspace
        }

        foreach ($scan in @($scanResult.Scans)) {
            $datasourcesById = @{}
            foreach ($source in @($scan.datasourceInstances) + @($scan.misconfiguredDatasourceInstances)) {
                if ($source.datasourceId) {
                    $datasourcesById[[string]$source.datasourceId] = $source
                }
            }

            foreach ($scannedWorkspace in @($scan.workspaces)) {
                $workspaceId = [string]$scannedWorkspace.id
                if ($workspaceById.ContainsKey($workspaceId)) {
                    $workspace = $workspaceById[$workspaceId]
                }
                else {
                    $workspace = $scannedWorkspace
                }

                foreach ($model in @($scannedWorkspace.datasets)) {
                    foreach ($usage in @($model.datasourceUsages) + @($model.misconfiguredDatasourceUsages)) {
                        $sourceId = [string]$usage.datasourceInstanceId
                        if (-not $sourceId -or -not $datasourcesById.ContainsKey($sourceId)) {
                            continue
                        }

                        $source = $datasourcesById[$sourceId]
                        $connection = $source.connectionDetails
                        if ($connection -is [string]) {
                            try { $connection = $connection | ConvertFrom-Json } catch { $connection = $null }
                        }

                        $server = $null
                        foreach ($propertyName in @('server', 'url', 'path', 'account', 'domain')) {
                            $value = $connection.$propertyName
                            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                                $server = [string]$value
                                break
                            }
                        }
                        if ([string]::IsNullOrWhiteSpace($server)) {
                            $server = 'Unknown endpoint'
                        }

                        $datasourceType = [string]$source.datasourceType
                        $category = Get-ServerCategory -DatasourceType $datasourceType
                        $normalizedServer = $server.Trim().TrimEnd('/').ToLowerInvariant()
                        $inventory.Add([PSCustomObject]@{
                            ServerKey = ("{0}|{1}" -f $category, $normalizedServer)
                            Server = $server
                            Category = $category
                            DatasourceType = $datasourceType
                            Database = [string]$connection.database
                            Workspace = $workspace.Name
                            WorkspaceId = [string]$workspaceId
                            SemanticModel = $model.name
                            SemanticModelId = [string]$model.id
                            DatasourceId = [string]$source.datasourceId
                            GatewayId = [string]$source.gatewayId
                        })
                    }
                }
            }
        }

        return [PSCustomObject]@{
            Items = $inventory.ToArray()
            FailedLookups = $scanResult.FailedBatches
            FirstLookupError = $scanResult.FirstError
        }
    }

    function Get-ServerSummaryRows {
        param ([object[]]$Inventory)

        @(
            foreach ($serverGroup in @($Inventory | Group-Object ServerKey)) {
                $first = $serverGroup.Group[0]
                [PSCustomObject]@{
                    ServerKey = $serverGroup.Name
                    Server = $first.Server
                    Category = $first.Category
                    DatasourceType = $first.DatasourceType
                    Datasources = $serverGroup.Count
                }
            }
        ) | Sort-Object Category, Server
    }

    function Show-ServerSummaryMenu {
        param (
            [object[]]$Inventory,
            [object[]]$ServerSummary,
            [object[]]$Reports,
            [int]$FailedLookups,
            [string]$FirstLookupError
        )

        $sqlCount = @($ServerSummary | Where-Object { $_.Category -eq 'SQL Server' }).Count
        $analysisServicesCount = @($ServerSummary | Where-Object { $_.Category -eq 'Analysis Services' }).Count
        $oracleCount = @($ServerSummary | Where-Object { $_.Category -eq 'Oracle' }).Count
        $otherCount = @($ServerSummary | Where-Object { $_.Category -eq 'Other' }).Count

        Write-Host "`nDATA SOURCE / SERVER SUMMARY"
        Write-Host ('-' * 60)
        Write-Host ("Total Servers              {0}" -f $ServerSummary.Count)
        Write-Host ("SQL Server                 {0}" -f $sqlCount)
        Write-Host ("Analysis Services          {0}" -f $analysisServicesCount)
        Write-Host ("Oracle                     {0}" -f $oracleCount)
        Write-Host ("Other                      {0}" -f $otherCount)
        if ($FailedLookups -gt 0) {
            Write-Host ("Semantic models not queried {0}" -f $FailedLookups)
            if ($FirstLookupError) {
                Write-Host ("First datasource error: {0}" -f $FirstLookupError)
            }
        }

        $back = $false
        while (-not $back) {
            Write-Host "`nSelect:"
            Write-Host "1. View all servers"
            Write-Host "2. Search server"
            Write-Host "3. Server -> Semantic Models"
            Write-Host "4. Server -> Workspaces"
            Write-Host "5. Server -> Reports"
            Write-Host "6. Export server inventory"
            Write-Host "7. Back"
            $choice = Read-Host 'Enter 1-7'

            switch ($choice) {
                '1' {
                    Write-DetailTable -Title 'SERVERS' -Items $ServerSummary
                }
                '2' {
                    $searchTerm = Read-Host 'Enter server name or text to search'
                    $matches = @($ServerSummary | Where-Object { $_.Server -like "*$searchTerm*" })
                    Write-DetailTable -Title 'SERVER SEARCH RESULTS' -Items $matches
                }
                '3' {
                    $rows = @($Inventory | Select-Object Server, Category, Workspace, SemanticModel, SemanticModelId | Sort-Object Server, Workspace, SemanticModel, SemanticModelId -Unique)
                    Write-DetailTable -Title 'SERVERS TO SEMANTIC MODELS' -Items $rows
                }
                '4' {
                    $rows = @($Inventory | Select-Object Server, Category, Workspace, WorkspaceId | Sort-Object Server, Workspace, WorkspaceId -Unique)
                    Write-DetailTable -Title 'SERVERS TO WORKSPACES' -Items $rows
                }
                '5' {
                    $rows = @(
                        foreach ($source in $Inventory) {
                            foreach ($report in @($Reports | Where-Object { $_.SemanticModel -eq $source.SemanticModelId })) {
                                [PSCustomObject]@{
                                    Server = $source.Server
                                    Category = $source.Category
                                    Workspace = $report.Workspace
                                    Report = $report.Name
                                    ReportId = $report.Id
                                    SemanticModel = $source.SemanticModel
                                }
                            }
                        }
                    ) | Sort-Object Server, Workspace, Report, ReportId -Unique
                    Write-DetailTable -Title 'SERVERS TO REPORTS' -Items $rows
                }
                '6' {
                    $defaultPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("PowerBI_ServerInventory_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                    $exportPath = Read-Host "CSV path (Enter for $defaultPath)"
                    if ([string]::IsNullOrWhiteSpace($exportPath)) {
                        $exportPath = $defaultPath
                    }
                    try {
                        $Inventory | Select-Object Server, Category, DatasourceType, Database, Workspace, WorkspaceId, SemanticModel, SemanticModelId, DatasourceId, GatewayId |
                            Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                        Write-Host "Server inventory exported to $exportPath"
                    }
                    catch {
                        Write-Error "Server inventory export failed: $_"
                    }
                }
                '7' { $back = $true }
                default { Write-Host 'Enter a number from 1 to 7.' }
            }
        }
    }

    function Invoke-PbiApi {
        param (
            [string]$Url,
            [ValidateSet('Get', 'Post', 'Put', 'Delete')]
            [string]$Method = 'Get',
            [string]$Body
        )

        $parameters = @{
            Url = $Url
            Method = $Method
            ErrorAction = 'Stop'
            WarningAction = 'SilentlyContinue'
        }
        if ($PSBoundParameters.ContainsKey('Body')) {
            $parameters.Body = $Body
            $parameters.ContentType = 'application/json'
        }

        $response = Invoke-PowerBIRestMethod @parameters
        if ($response -is [string]) {
            return ($response | ConvertFrom-Json)
        }
        return $response
    }

    function Read-ConsoleMenu {
        param (
            [string]$Title,
            [string[]]$Options
        )

        do {
            Write-Host "`n$Title"
            for ($optionIndex = 0; $optionIndex -lt $Options.Count; $optionIndex++) {
                Write-Host ("{0}. {1}" -f ($optionIndex + 1), $Options[$optionIndex])
            }
            $choice = Read-Host ("Choose 1-{0}" -f $Options.Count)
            if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $Options.Count) {
                Write-Host ("Enter a number from 1 to {0}." -f $Options.Count)
                $choice = $null
            }
        } while (-not $choice)

        return [int]$choice
    }

    function Export-ConsoleRows {
        param (
            [object[]]$Rows,
            [string]$BaseName
        )

        $defaultPath = Join-Path $script:OutputFolder ("{0}_{1}.csv" -f $BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $path = Read-Host "CSV path (Enter for $defaultPath)"
        if ([string]::IsNullOrWhiteSpace($path)) {
            $path = $defaultPath
        }
        try {
            $parentFolder = Split-Path -Path $path -Parent
            if ($parentFolder -and -not (Test-Path -LiteralPath $parentFolder -PathType Container)) {
                New-Item -ItemType Directory -Path $parentFolder -Force -ErrorAction Stop | Out-Null
            }
            $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "Exported: $path"
        }
        catch {
            Write-Error "Export failed: $_ Check destination permissions and Windows Security Controlled Folder Access. If this folder is protected, contact IT for an approved export location."
        }
    }

    function Get-WorkspaceCapacityName {
        param ([object]$Workspace)

        if ($Workspace.CapacityId) {
            $capacity = $script:Capacities | Where-Object { $_.Id -eq $Workspace.CapacityId } | Select-Object -First 1
            if ($capacity) { return $capacity.DisplayName }
            return [string]$Workspace.CapacityId
        }
        return 'Shared'
    }

    function Get-ScanMetadata {
        if ($null -ne $script:ScanWorkspaces) {
            return $script:ScanWorkspaces
        }

        $script:ScanWorkspaces = @()
        $workspaceIds = @($script:Workspaces | ForEach-Object { [string]$_.Id } | Where-Object { $_ })
        for ($offset = 0; $offset -lt $workspaceIds.Count; $offset += 100) {
            $batch = @($workspaceIds | Select-Object -Skip $offset -First 100)
            $body = @{ workspaces = $batch } | ConvertTo-Json -Depth 5
            $scan = Invoke-PbiApi -Url 'admin/workspaces/getInfo?lineage=true&datasourceDetails=true&datasetSchema=true&datasetExpressions=true&getArtifactUsers=true' -Method Post -Body $body
            $scanId = [string]$scan.id
            if (-not $scanId) { continue }

            $status = 'NotStarted'
            for ($attempt = 0; $attempt -lt 60 -and $status -notin @('Succeeded', 'Failed'); $attempt++) {
                Start-Sleep -Seconds 2
                $scanStatus = Invoke-PbiApi -Url ("admin/workspaces/scanStatus/{0}" -f $scanId)
                $status = [string]$scanStatus.status
            }
            if ($status -ne 'Succeeded') {
                Write-Host ("Metadata scan {0} ended with status {1}." -f $scanId, $status)
                continue
            }

            $result = Invoke-PbiApi -Url ("admin/workspaces/scanResult/{0}" -f $scanId)
            $script:ScanWorkspaces += @($result.workspaces)
        }

        return $script:ScanWorkspaces
    }

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

    function Get-GatewayInventory {
        if ($null -ne $script:GatewayInventory) {
            return $script:GatewayInventory
        }

        $script:GatewayInventory = @()
        $gatewayResult = Invoke-PbiApi -Url 'gateways'
        foreach ($gateway in @($gatewayResult.value)) {
            $sources = @()
            try {
                $sourceResult = Invoke-PbiApi -Url ("gateways/{0}/datasources" -f $gateway.Id)
                $sources = @($sourceResult.value)
            }
            catch {
                Write-Host ("Could not read datasource list for gateway {0}: {1}" -f $gateway.Name, $_.Exception.Message)
            }

            foreach ($source in $sources) {
                $connection = $null
                if ($source.ConnectionDetails) {
                    try { $connection = $source.ConnectionDetails | ConvertFrom-Json } catch { $connection = $null }
                }
                $script:GatewayInventory += [PSCustomObject]@{
                    Gateway = $gateway.Name
                    GatewayId = [string]$gateway.Id
                    GatewayStatus = [string]$gateway.GatewayStatus
                    Datasource = [string]$source.DatasourceName
                    DatasourceId = [string]$source.Id
                    DatasourceType = [string]$source.DatasourceType
                    Server = [string]$connection.server
                    Database = [string]$connection.database
                    ConnectionDetails = [string]$source.ConnectionDetails
                }
            }
            if ($sources.Count -eq 0) {
                $script:GatewayInventory += [PSCustomObject]@{
                    Gateway = $gateway.Name
                    GatewayId = [string]$gateway.Id
                    GatewayStatus = [string]$gateway.GatewayStatus
                    Datasource = $null
                    DatasourceId = $null
                    DatasourceType = $null
                    Server = $null
                    Database = $null
                    ConnectionDetails = $null
                }
            }
        }
        return $script:GatewayInventory
    }

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

    $script:OutputFolder = $PSScriptRoot
    $script:Workspaces = @($workspaces)
    $script:PublicWorkspaces = @($publicWorkspaces)
    $script:PersonalWorkspaces = @($personalWorkspaces)
    $script:DeletedWorkspaces = @($deletedWorkspaces)
    $script:OrphanedWorkspaces = @($orphanedWorkspaces)
    $script:Capacities = @()
    $script:ScanWorkspaces = $null
    $script:RefreshInventory = $null
    $script:GatewayInventory = $null
    $script:ActivityInventory = $null
    $script:ServerInventory = $null
    $script:ServerSummaryRows = @()
    $script:ServerLookupFailures = 0

    function Get-ConsoleContent {
        param ([object[]]$WorkspaceSet)

        $reports = @(
            foreach ($workspace in $WorkspaceSet) {
                foreach ($report in @($workspace.Reports)) {
                    [PSCustomObject]@{
                        Name = $report.Name
                        Id = [string]$report.Id
                        Workspace = $workspace.Name
                        WorkspaceId = [string]$workspace.Id
                        SemanticModel = [string]$report.DatasetId
                        Type = Get-ReportKind -Report $report
                        WebUrl = [string]$report.WebUrl
                    }
                }
            }
        ) | Sort-Object Workspace, Name
        $models = @(
            foreach ($workspace in $WorkspaceSet) {
                foreach ($model in @($workspace.Datasets)) {
                    [PSCustomObject]@{
                        Name = $model.Name
                        Id = [string]$model.Id
                        Workspace = $workspace.Name
                        WorkspaceId = [string]$workspace.Id
                        Owner = $model.ConfiguredBy
                        IsRefreshable = $model.IsRefreshable
                    }
                }
            }
        ) | Sort-Object Workspace, Name
        $dashboards = @(
            foreach ($workspace in $WorkspaceSet) {
                foreach ($dashboard in @($workspace.Dashboards)) {
                    [PSCustomObject]@{ Name = Get-ArtifactName -Item $dashboard; Id = [string]$dashboard.Id; Workspace = $workspace.Name; WorkspaceId = [string]$workspace.Id }
                }
            }
        ) | Sort-Object Workspace, Name
        $dataflows = @(
            foreach ($workspace in $WorkspaceSet) {
                foreach ($dataflow in @($workspace.Dataflows)) {
                    [PSCustomObject]@{ Name = $dataflow.Name; Id = [string]$dataflow.ObjectId; Workspace = $workspace.Name; WorkspaceId = [string]$workspace.Id; Owner = $dataflow.ConfiguredBy }
                }
            }
        ) | Sort-Object Workspace, Name

        return [PSCustomObject]@{ Reports = @($reports); Models = @($models); Dashboards = @($dashboards); Dataflows = @($dataflows) }
    }

    function Select-WorkspaceScope {
        $choice = Read-ConsoleMenu -Title 'Workspace scope' -Options @('Public Workspaces', 'Personal Workspaces', 'Both', 'Back')
        switch ($choice) {
            1 { return @($script:PublicWorkspaces) }
            2 { return @($script:PersonalWorkspaces) }
            3 { return @($script:PublicWorkspaces + $script:PersonalWorkspaces | Sort-Object Id -Unique) }
            default { return @() }
        }
    }

    function Show-TenantSummary {
        $choice = Read-ConsoleMenu -Title 'TENANT SUMMARY' -Options @('Public Workspaces', 'Personal Workspaces', 'Both', 'Capacity Summary', 'Back')
        switch ($choice) {
            1 {
                $active = @($script:PublicWorkspaces | Where-Object State -eq 'Active').Count
                $deleted = @($script:DeletedWorkspaces | Where-Object Type -in @('Workspace', 'Group')).Count
                $orphaned = @($script:OrphanedWorkspaces | Where-Object Type -in @('Workspace', 'Group')).Count
                Write-Host "`nPUBLIC WORKSPACES"
                Write-Host ("Total       {0}" -f $script:PublicWorkspaces.Count)
                Write-Host ("Active      {0}" -f $active)
                Write-Host ("Deleted     {0}" -f $deleted)
                Write-Host ("Orphaned    {0}" -f $orphaned)
            }
            2 {
                $active = @($script:PersonalWorkspaces | Where-Object State -eq 'Active').Count
                Write-Host "`nPERSONAL WORKSPACES"
                Write-Host ("Total       {0}" -f $script:PersonalWorkspaces.Count)
                Write-Host ("Active      {0}" -f $active)
                Write-Host ("Inactive    {0}" -f ($script:PersonalWorkspaces.Count - $active))
            }
            3 {
                Write-WorkspaceSummary -Title 'PUBLIC WORKSPACES' -CurrentWorkspaces $script:PublicWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPublic
                Write-WorkspaceSummary -Title 'PERSONAL WORKSPACES' -CurrentWorkspaces $script:PersonalWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPersonal
            }
            4 {
                try {
                    if ($script:Capacities.Count -eq 0) {
                        $capacityResponse = Invoke-PbiApi -Url 'capacities'
                        $script:Capacities = @($capacityResponse.value)
                    }
                    Write-DetailTable -Title 'CAPACITIES' -Items @($script:Capacities | Select-Object DisplayName, Id, Sku, State, Region, CapacityUserAccessRight)
                }
                catch { Write-Error "Capacity list unavailable: $_" }
            }
        }
    }

    function Show-WorkspacesSection {
        $choice = Read-ConsoleMenu -Title 'WORKSPACES' -Options @('Public Workspaces', 'Personal Workspaces', 'Search Workspace', 'Workspace Details', 'Deleted Workspaces', 'Orphaned Workspaces', 'Export Workspace Inventory', 'Back')
        switch ($choice) {
            { $_ -in 1, 2 } {
                $scope = if ($choice -eq 1) { $script:PublicWorkspaces } else { $script:PersonalWorkspaces }
                Write-DetailTable -Title 'WORKSPACES' -Items @($scope | Select-Object Name, Id, State, Type, CapacityId)
            }
            3 {
                $term = Read-Host 'Workspace name or ID'
                Write-DetailTable -Title 'WORKSPACE SEARCH' -Items @($script:Workspaces | Where-Object { $_.Name -like "*$term*" -or $_.Id -like "*$term*" } | Select-Object Name, Id, State, Type, CapacityId)
            }
            4 {
                $term = Read-Host 'Workspace name or ID'
                $matches = @($script:Workspaces | Where-Object { $_.Name -like "*$term*" -or $_.Id -eq $term })
                $details = @(
                    foreach ($workspace in $matches) {
                        [PSCustomObject]@{
                            Name = $workspace.Name
                            Id = $workspace.Id
                            State = $workspace.State
                            Type = $workspace.Type
                            Capacity = Get-WorkspaceCapacityName -Workspace $workspace
                            Reports = @($workspace.Reports).Count
                            SemanticModels = @($workspace.Datasets).Count
                            Dataflows = @($workspace.Dataflows).Count
                            Dashboards = @($workspace.Dashboards).Count
                        }
                    }
                )
                Write-DetailTable -Title 'WORKSPACE DETAILS' -Items $details
                foreach ($workspace in $matches) {
                    try {
                        $users = Invoke-PbiApi -Url "groups/$($workspace.Id)/users"
                        Write-DetailTable -Title ("USERS: {0}" -f $workspace.Name) -Items @($users.value | Select-Object DisplayName, EmailAddress, Identifier, GroupUserAccessRight, PrincipalType)
                    }
                    catch { Write-Host ("Workspace users unavailable for {0}: {1}" -f $workspace.Name, $_.Exception.Message) }
                }
            }
            5 { Write-DetailTable -Title 'DELETED WORKSPACES' -Items @($script:DeletedWorkspaces | Select-Object Name, Id, State, Type) }
            6 { Write-DetailTable -Title 'ORPHANED WORKSPACES' -Items @($script:OrphanedWorkspaces | Select-Object Name, Id, State, Type) }
            7 {
                $rows = @($script:Workspaces | Select-Object Name, Id, State, Type, CapacityId)
                Export-ConsoleRows -Rows $rows -BaseName 'PowerBI_Workspaces'
            }
        }
    }

    function Show-ContentSection {
        $choice = Read-ConsoleMenu -Title 'CONTENT' -Options @('Power BI Reports', 'Paginated Reports', 'Semantic Models', 'Dashboards', 'Dataflows', 'Search Content', 'Export Content Inventory', 'Back')
        $scope = Select-WorkspaceScope
        if ($scope.Count -eq 0) { return }
        $content = Get-ConsoleContent -WorkspaceSet $scope
        switch ($choice) {
            1 { Write-DetailTable -Title 'POWER BI REPORTS' -Items @($content.Reports | Where-Object Type -eq 'PowerBIReport' | Select-Object Name, Id, Workspace, SemanticModel, Type, WebUrl) }
            2 { Write-DetailTable -Title 'PAGINATED REPORTS' -Items @($content.Reports | Where-Object Type -eq 'PaginatedReport' | Select-Object Name, Workspace, Id, SemanticModel) }
            3 {
                $metadata = Get-ScanMetadata
                $models = @($metadata | ForEach-Object { $workspace = $_; foreach ($model in @($_.Datasets)) { [PSCustomObject]@{ Name = $model.Name; Id = $model.Id; Workspace = $workspace.Name; Owner = $model.ConfiguredBy; Refreshable = $model.IsRefreshable; Tables = @($model.Tables).Count; Columns = (@($model.Tables | ForEach-Object { @($_.Columns).Count }) | Measure-Object -Sum).Sum; Measures = (@($model.Tables | ForEach-Object { @($_.Measures).Count }) | Measure-Object -Sum).Sum } } })
                Write-DetailTable -Title 'SEMANTIC MODELS' -Items $models
            }
            4 { Write-DetailTable -Title 'DASHBOARDS' -Items $content.Dashboards }
            5 { Write-DetailTable -Title 'DATAFLOWS' -Items $content.Dataflows }
            6 {
                $term = Read-Host 'Search reports, semantic models, dashboards, or dataflows'
                $matches = @($content.Reports + $content.Models + $content.Dashboards + $content.Dataflows | Where-Object { $_.Name -like "*$term*" -or $_.Workspace -like "*$term*" -or $_.Id -like "*$term*" })
                Write-DetailTable -Title 'CONTENT SEARCH RESULTS' -Items $matches
            }
            7 { Export-ConsoleRows -Rows @($content.Reports + $content.Models + $content.Dashboards + $content.Dataflows) -BaseName 'PowerBI_Content' }
        }
    }

    function Show-DataSourcesSection {
        if ($null -eq $script:ServerInventory) {
            Write-Host 'Retrieving datasource details for semantic models...'
            $inventoryResult = Get-ServerInventory -Workspaces $script:Workspaces
            $script:ServerInventory = @($inventoryResult.Items)
            $script:ServerLookupFailures = $inventoryResult.FailedLookups
            $script:ServerLookupError = $inventoryResult.FirstLookupError
            $script:ServerSummaryRows = @(Get-ServerSummaryRows -Inventory $script:ServerInventory)
        }

        $choice = Read-ConsoleMenu -Title 'DATA SOURCES & SERVERS' -Options @('Server Summary', 'View All Servers', 'Search Server', 'Server Impact Analysis', 'Database Summary', 'Datasource Types', 'Gateways', 'Export Server Inventory', 'Back')
        switch ($choice) {
            1 { Show-ServerSummaryMenu -Inventory $script:ServerInventory -ServerSummary $script:ServerSummaryRows -Reports @((Get-ConsoleContent -WorkspaceSet $script:Workspaces).Reports) -FailedLookups $script:ServerLookupFailures -FirstLookupError $script:ServerLookupError }
            2 { Write-DetailTable -Title 'SERVERS' -Items $script:ServerSummaryRows }
            3 {
                $term = Read-Host 'Server name or text'
                Write-DetailTable -Title 'SERVER SEARCH' -Items @($script:ServerSummaryRows | Where-Object Server -like "*$term*")
            }
            4 {
                $term = Read-Host 'Server name or text'
                $matches = @($script:ServerInventory | Where-Object Server -like "*$term*")
                Write-DetailTable -Title 'SERVER IMPACT ANALYSIS' -Items @($matches | Select-Object Server, Category, Database, Workspace, SemanticModel, SemanticModelId | Sort-Object Workspace, SemanticModel -Unique)
            }
            5 { Write-DetailTable -Title 'DATABASE SUMMARY' -Items @($script:ServerInventory | Group-Object Server, Database, DatasourceType | ForEach-Object { [PSCustomObject]@{ Server = $_.Group[0].Server; Database = $_.Group[0].Database; DatasourceType = $_.Group[0].DatasourceType; SemanticModels = @($_.Group.SemanticModelId | Sort-Object -Unique).Count; Workspaces = @($_.Group.WorkspaceId | Sort-Object -Unique).Count } }) }
            6 { Write-DetailTable -Title 'DATASOURCE TYPES' -Items @($script:ServerInventory | Group-Object DatasourceType | Select-Object @{Name='DatasourceType';Expression={$_.Name}}, Count) }
            7 { Show-GatewaySection }
            8 { Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_ServerInventory' }
        }
    }

    function Show-QuerySection {
        $choice = Read-ConsoleMenu -Title 'QUERIES & EXPRESSIONS' -Options @('Search Query', 'Queries by Semantic Model', 'Queries by Server', 'Queries by Database', 'Power Query / Mashup Expressions', 'DAX Expressions', 'Tables & Columns', 'Search Query Text', 'Export Query Inventory', 'Back')
        $metadata = Get-ScanMetadata
        $rows = @(
            foreach ($workspace in $metadata) {
                foreach ($model in @($workspace.Datasets)) {
                    foreach ($table in @($model.Tables)) {
                        foreach ($source in @($table.Source)) {
                            [PSCustomObject]@{ Workspace = $workspace.Name; SemanticModel = $model.Name; SemanticModelId = $model.Id; Table = $table.Name; Query = $source.Expression }
                        }
                    }
                    foreach ($expression in @($model.Expressions)) {
                        [PSCustomObject]@{ Workspace = $workspace.Name; SemanticModel = $model.Name; SemanticModelId = $model.Id; Table = $null; Query = $expression.Expression }
                    }
                    foreach ($table in @($model.Tables)) {
                        foreach ($measure in @($table.Measures)) {
                            [PSCustomObject]@{ Workspace = $workspace.Name; SemanticModel = $model.Name; SemanticModelId = $model.Id; Table = $table.Name; Query = $measure.Expression; Kind = 'Measure' }
                        }
                    }
                }
            }
        )
        switch ($choice) {
            1 { $term = Read-Host 'Query text'; Write-DetailTable -Title 'QUERY SEARCH' -Items @($rows | Where-Object Query -like "*$term*") }
            2 { $term = Read-Host 'Semantic model name'; Write-DetailTable -Title 'QUERIES BY SEMANTIC MODEL' -Items @($rows | Where-Object SemanticModel -like "*$term*") }
            3 { $term = Read-Host 'Server name'; Write-DetailTable -Title 'QUERIES BY SERVER' -Items @($rows | Where-Object { $_.Query -like "*$term*" }) }
            4 { $term = Read-Host 'Database name'; Write-DetailTable -Title 'QUERIES BY DATABASE' -Items @($rows | Where-Object { $_.Query -like "*$term*" }) -Properties @('Workspace', 'SemanticModel', 'SemanticModelId', 'Query') }
            5 { Write-DetailTable -Title 'POWER QUERY / MASHUP EXPRESSIONS' -Items @($rows | Where-Object { $_.Kind -ne 'Measure' }) }
            6 { Write-DetailTable -Title 'DAX EXPRESSIONS' -Items @($rows | Where-Object Kind -eq 'Measure') }
            7 { Write-DetailTable -Title 'TABLES & COLUMNS' -Items @($metadata | ForEach-Object { $workspace = $_; foreach ($model in @($_.Datasets)) { foreach ($table in @($model.Tables)) { foreach ($column in @($table.Columns)) { [PSCustomObject]@{ Workspace=$workspace.Name; SemanticModel=$model.Name; Table=$table.Name; Column=$column.Name; DataType=$column.DataType; Hidden=$column.IsHidden } } } } }) }
            8 { $term = Read-Host 'Text'; Write-DetailTable -Title 'QUERY TEXT SEARCH' -Items @($rows | Where-Object Query -like "*$term*") }
            9 { Export-ConsoleRows -Rows $rows -BaseName 'PowerBI_QueryInventory' }
        }
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

    function Show-GatewaySection {
        $choice = Read-ConsoleMenu -Title 'GATEWAY MONITORING' -Options @('All Gateways', 'Gateway Datasources', 'Server to Gateway Mapping', 'Semantic Model to Gateway Mapping', 'Search Gateway', 'Export Gateway Inventory', 'Back')
        $gateways = @(Get-GatewayInventory)
        switch ($choice) {
            1 { Write-DetailTable -Title 'GATEWAYS' -Items @($gateways | Select-Object Gateway, GatewayId, GatewayStatus | Sort-Object Gateway -Unique) }
            2 { Write-DetailTable -Title 'GATEWAY DATASOURCES' -Items $gateways }
            3 { Write-DetailTable -Title 'SERVER TO GATEWAY' -Items @($gateways | Select-Object Server, Database, DatasourceType, Gateway, GatewayId | Sort-Object Server, Database -Unique) }
            4 { Write-DetailTable -Title 'MODEL TO GATEWAY' -Items @($gateways | ForEach-Object { $row = $_; foreach ($source in $script:ServerInventory | Where-Object { $_.Server -eq $row.Server -and $_.Database -eq $row.Database }) { [PSCustomObject]@{ Workspace=$source.Workspace; SemanticModel=$source.SemanticModel; Server=$row.Server; Gateway=$row.Gateway; GatewayId=$row.GatewayId } } } | Sort-Object Workspace, SemanticModel -Unique) }
            5 { $term = Read-Host 'Gateway name, server, or datasource'; Write-DetailTable -Title 'GATEWAY SEARCH' -Items @($gateways | Where-Object { $_.Gateway -like "*$term*" -or $_.Server -like "*$term*" -or $_.Datasource -like "*$term*" }) }
            6 { Export-ConsoleRows -Rows $gateways -BaseName 'PowerBI_GatewayInventory' }
        }
    }

    function Show-SecuritySection {
        $choice = Read-ConsoleMenu -Title 'SECURITY & ACCESS' -Options @('Workspace Users', 'Search User', 'Workspace Admins', 'Orphaned Workspaces', 'Export Permissions', 'Back')
        $userRows = @(
            foreach ($workspace in $script:Workspaces) {
                try {
                    $response = Invoke-PbiApi -Url "groups/$($workspace.Id)/users"
                    foreach ($user in @($response.value)) {
                        [PSCustomObject]@{ Workspace=$workspace.Name; WorkspaceId=[string]$workspace.Id; DisplayName=$user.DisplayName; EmailAddress=$user.EmailAddress; Identifier=$user.Identifier; Access=$user.GroupUserAccessRight; PrincipalType=$user.PrincipalType }
                    }
                }
                catch { }
            }
        )
        switch ($choice) {
            1 { Write-DetailTable -Title 'WORKSPACE USERS' -Items $userRows }
            2 { $term = Read-Host 'User name or email'; Write-DetailTable -Title 'USER SEARCH' -Items @($userRows | Where-Object { $_.DisplayName -like "*$term*" -or $_.EmailAddress -like "*$term*" -or $_.Identifier -like "*$term*" }) }
            3 { Write-DetailTable -Title 'WORKSPACE ADMINS' -Items @($userRows | Where-Object Access -eq 'Admin') }
            4 { Write-DetailTable -Title 'ORPHANED WORKSPACES' -Items @($script:OrphanedWorkspaces | Select-Object Name, Id, Type, State) }
            5 { Export-ConsoleRows -Rows $userRows -BaseName 'PowerBI_Permissions' }
        }
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

    function Show-ExportSection {
        $choice = Read-ConsoleMenu -Title 'EXPORT & REPORTING' -Options @('Workspace Inventory', 'Report Inventory', 'Semantic Model Inventory', 'Server Inventory', 'Database Inventory', 'Datasource Inventory', 'Query Inventory', 'Refresh Failures', 'Permissions', 'Activity', 'Complete Tenant Inventory', 'Back')
        switch ($choice) {
            1 { Export-ConsoleRows -Rows @($script:Workspaces | Select-Object Name, Id, Type, State, CapacityId) -BaseName 'PowerBI_Workspaces' }
            2 { Export-ConsoleRows -Rows @(Get-ConsoleContent -WorkspaceSet $script:Workspaces).Reports -BaseName 'PowerBI_Reports' }
            3 { Export-ConsoleRows -Rows @(Get-ConsoleContent -WorkspaceSet $script:Workspaces).Models -BaseName 'PowerBI_SemanticModels' }
            4 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_Servers' }
            5 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows @($script:ServerInventory | Select-Object Server, Database, DatasourceType | Sort-Object Server, Database, DatasourceType -Unique) -BaseName 'PowerBI_Databases' }
            6 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_Datasources' }
            7 { $metadata = Get-ScanMetadata; Export-ConsoleRows -Rows @($metadata | ForEach-Object { $workspace=$_; foreach ($model in @($_.Datasets)) { foreach ($expression in @($model.Expressions)) { [PSCustomObject]@{ Workspace=$workspace.Name; SemanticModel=$model.Name; Expression=$expression.Name; Query=$expression.Expression } } } }) -BaseName 'PowerBI_Queries' }
            8 { Export-ConsoleRows -Rows @((Get-RefreshInventory) | Where-Object Status -eq 'Failed') -BaseName 'PowerBI_RefreshFailures' }
            9 { $users = @(); foreach ($workspace in $script:Workspaces) { try { $r=Invoke-PbiApi -Url "groups/$($workspace.Id)/users"; $users += @($r.value | Select-Object @{Name='Workspace';Expression={$workspace.Name}}, DisplayName, EmailAddress, Identifier, GroupUserAccessRight) } catch { } }; Export-ConsoleRows -Rows $users -BaseName 'PowerBI_Permissions' }
            10 { Export-ConsoleRows -Rows (Get-ActivityInventory) -BaseName 'PowerBI_Activity' }
            11 {
                $content = Get-ConsoleContent -WorkspaceSet $script:Workspaces
                $rows = @($script:Workspaces | Select-Object @{Name='RecordType';Expression={'Workspace'}}, Name, Id, Type, State) + @($content.Reports | Select-Object @{Name='RecordType';Expression={'Report'}}, Name, Id, Type, Workspace) + @($content.Models | Select-Object @{Name='RecordType';Expression={'SemanticModel'}}, Name, Id, Owner, Workspace) + @($content.Dashboards | Select-Object @{Name='RecordType';Expression={'Dashboard'}}, Name, Id, Workspace) + @($content.Dataflows | Select-Object @{Name='RecordType';Expression={'Dataflow'}}, Name, Id, Owner, Workspace)
                Export-ConsoleRows -Rows $rows -BaseName 'PowerBI_CompleteTenantInventory'
            }
        }
    }

    function Show-SettingsSection {
        $choice = Read-ConsoleMenu -Title 'SETTINGS' -Options @('Output Folder', 'Clear Cache', 'Authentication Status', 'Back')
        switch ($choice) {
            1 {
                Write-Host ("Current output folder: {0}" -f $script:OutputFolder)
                $folder = Read-Host 'New output folder (blank to keep current)'
                if (-not [string]::IsNullOrWhiteSpace($folder)) {
                    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
                    $script:OutputFolder = (Resolve-Path -LiteralPath $folder).Path
                }
            }
            2 { $script:ScanWorkspaces = $null; $script:RefreshInventory = $null; $script:GatewayInventory = $null; $script:ActivityInventory = $null; $script:ServerInventory = $null; Write-Host 'Cached inventories cleared.' }
            3 { $profile = Get-PowerBIProfile; Write-Host ("Signed in user: {0}" -f $profile.UserName) }
        }
    }

    function Show-QueriesSection {
        $choice = Read-ConsoleMenu -Title 'QUERIES & EXPRESSIONS' -Options @('Search Query', 'Queries by Semantic Model', 'Queries by Server', 'Queries by Database', 'Power Query / Mashup Expressions', 'DAX Expressions', 'Tables & Columns', 'Search Query Text', 'Export Query Inventory', 'Back')
        $metadata = Get-ScanMetadata
        $queries = @(
            foreach ($workspace in $metadata) {
                foreach ($model in @($workspace.Datasets)) {
                    foreach ($expression in @($model.Expressions)) { [PSCustomObject]@{ Workspace=$workspace.Name; SemanticModel=$model.Name; SemanticModelId=$model.Id; Table=$null; Expression=$expression.Name; Query=$expression.Expression; Kind='Expression' } }
                    foreach ($table in @($model.Tables)) {
                        foreach ($source in @($table.Source)) { [PSCustomObject]@{ Workspace=$workspace.Name; SemanticModel=$model.Name; SemanticModelId=$model.Id; Table=$table.Name; Expression=$null; Query=$source.Expression; Kind='Mashup' } }
                        foreach ($measure in @($table.Measures)) { [PSCustomObject]@{ Workspace=$workspace.Name; SemanticModel=$model.Name; SemanticModelId=$model.Id; Table=$table.Name; Expression=$measure.Name; Query=$measure.Expression; Kind='Measure' } }
                    }
                }
            }
        )
        switch ($choice) {
            1 { $term=Read-Host 'Query text'; Write-DetailTable -Title 'QUERY SEARCH' -Items @($queries | Where-Object Query -like "*$term*") }
            2 { $term=Read-Host 'Semantic model'; Write-DetailTable -Title 'QUERIES BY MODEL' -Items @($queries | Where-Object SemanticModel -like "*$term*") }
            3 { $term=Read-Host 'Server'; Write-DetailTable -Title 'QUERIES BY SERVER' -Items @($queries | Where-Object Query -like "*$term*") }
            4 { $term=Read-Host 'Database'; Write-DetailTable -Title 'QUERIES BY DATABASE' -Items @($queries | Where-Object Query -like "*$term*") -Properties @('Workspace', 'SemanticModel', 'SemanticModelId', 'Query') }
            5 { Write-DetailTable -Title 'POWER QUERY / MASHUP' -Items @($queries | Where-Object Kind -in @('Mashup','Expression')) }
            6 { Write-DetailTable -Title 'DAX EXPRESSIONS' -Items @($queries | Where-Object Kind -eq 'Measure') }
            7 { Write-DetailTable -Title 'TABLES & COLUMNS' -Items @($metadata | ForEach-Object { $workspace=$_; foreach ($model in @($_.Datasets)) { foreach ($table in @($model.Tables)) { foreach ($column in @($table.Columns)) { [PSCustomObject]@{Workspace=$workspace.Name; SemanticModel=$model.Name; Table=$table.Name; Column=$column.Name; DataType=$column.DataType; Hidden=$column.IsHidden} } } } }) }
            8 { $term=Read-Host 'Text to search'; Write-DetailTable -Title 'QUERY TEXT SEARCH' -Items @($queries | Where-Object Query -like "*$term*") }
            9 { Export-ConsoleRows -Rows $queries -BaseName 'PowerBI_QueryInventory' }
        }
    }

    function Show-ImpactSection {
        $choice = Read-ConsoleMenu -Title 'IMPACT ANALYSIS' -Options @('Search Server', 'Search Database', 'Search Semantic Model', 'Search Report', 'Export Impact Analysis', 'Back')
        if ($null -eq $script:ServerInventory) { $r=Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory=@($r.Items) }
        $content = Get-ConsoleContent -WorkspaceSet $script:Workspaces
        switch ($choice) {
            { $_ -in 1, 2 } {
                $term = Read-Host $(if ($choice -eq 1) { 'Server' } else { 'Database' })
                $sources = @($script:ServerInventory | Where-Object { if ($choice -eq 1) { $_.Server -like "*$term*" } else { $_.Database -like "*$term*" } })
                $modelIds = @($sources.SemanticModelId | Sort-Object -Unique)
                $rows = @($content.Reports | Where-Object { $_.SemanticModel -in $modelIds } | Select-Object Name, Id, Workspace, SemanticModel, @{Name='Server';Expression={($sources | Where-Object SemanticModelId -eq $_.SemanticModel | Select-Object -ExpandProperty Server -Unique) -join '; '}})
                Write-DetailTable -Title 'IMPACT ANALYSIS' -Items $rows
            }
            3 { $term=Read-Host 'Semantic model'; Write-DetailTable -Title 'MODEL IMPACT' -Items @($content.Reports | Where-Object SemanticModel -like "*$term*" | Select-Object Name, Id, Workspace, SemanticModel) }
            4 { $term=Read-Host 'Report'; Write-DetailTable -Title 'REPORT IMPACT' -Items @($content.Reports | Where-Object Name -like "*$term*" | Select-Object Name, Id, Workspace, SemanticModel) }
            5 { Export-ConsoleRows -Rows @($content.Reports | Select-Object Name, Id, Workspace, SemanticModel) -BaseName 'PowerBI_ImpactAnalysis' }
        }
    }

    function Show-GlobalSearch {
        $choice = Read-ConsoleMenu -Title 'GLOBAL SEARCH' -Options @('Search Server', 'Search Database', 'Search Workspace', 'Search Report', 'Search Semantic Model', 'Search User', 'Search Query / Expression', 'Back')
        $term = Read-Host 'Search text'
        switch ($choice) {
            1 { if ($null -eq $script:ServerInventory) { $r=Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory=@($r.Items) }; Write-DetailTable -Title 'SERVER RESULTS' -Items @($script:ServerInventory | Where-Object Server -like "*$term*") }
            2 { if ($null -eq $script:ServerInventory) { $r=Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory=@($r.Items) }; Write-DetailTable -Title 'DATABASE RESULTS' -Items @($script:ServerInventory | Where-Object Database -like "*$term*") }
            3 { Write-DetailTable -Title 'WORKSPACE RESULTS' -Items @($script:Workspaces | Where-Object { $_.Name -like "*$term*" -or $_.Id -like "*$term*" } | Select-Object Name, Id, Type, State) }
            4 { $c=Get-ConsoleContent -WorkspaceSet $script:Workspaces; Write-DetailTable -Title 'REPORT RESULTS' -Items @($c.Reports | Where-Object { $_.Name -like "*$term*" -or $_.Id -like "*$term*" }) }
            5 { $c=Get-ConsoleContent -WorkspaceSet $script:Workspaces; Write-DetailTable -Title 'SEMANTIC MODEL RESULTS' -Items @($c.Models | Where-Object { $_.Name -like "*$term*" -or $_.Id -like "*$term*" }) }
            6 { Show-SecuritySection }
            7 { $m=Get-ScanMetadata; Write-DetailTable -Title 'QUERY RESULTS' -Items @($m | ForEach-Object { $workspace=$_; foreach ($model in @($_.Datasets)) { foreach ($expression in @($model.Expressions)) { if ($expression.Expression -like "*$term*") { [PSCustomObject]@{Workspace=$workspace.Name; SemanticModel=$model.Name; Expression=$expression.Name; Query=$expression.Expression} } } } }) }
        }
    }

    $mainOptions = @(
        'TENANT SUMMARY', 'WORKSPACES', 'CONTENT', 'DATA SOURCES & SERVERS',
        'QUERIES & EXPRESSIONS', 'REFRESH MONITORING', 'GATEWAY MONITORING',
        'SECURITY & ACCESS', 'ACTIVITY MONITORING', 'IMPACT ANALYSIS',
        'GLOBAL SEARCH', 'EXPORT & REPORTING', 'SETTINGS', 'EXIT'
    )
    $running = $true
    while ($running) {
        Write-Host "`n============================================================"
        Write-Host '              POWER BI SERVICE ADMIN CONSOLE'
        Write-Host "============================================================"
        $mainChoice = Read-ConsoleMenu -Title 'MAIN MENU' -Options $mainOptions
        try {
            switch ($mainChoice) {
                1 { Show-TenantSummary }
                2 { Show-WorkspacesSection }
                3 { Show-ContentSection }
                4 { Show-DataSourcesSection }
                5 { Show-QueriesSection }
                6 { Show-RefreshSection }
                7 { Show-GatewaySection }
                8 { Show-SecuritySection }
                9 { Show-ActivitySection }
                10 { Show-ImpactSection }
                11 { Show-GlobalSearch }
                12 { Show-ExportSection }
                13 { Show-SettingsSection }
                14 { $running = $false }
            }
        }
        catch {
            Write-Error ("Section failed: {0}" -f $_.Exception.Message)
        }
    }
}
catch {
    Write-Error "Power BI sign-in or summary retrieval failed: $_"
    exit 1
}