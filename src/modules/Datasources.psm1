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
                $defaultPath = Join-Path $script:OutputFolder 'PowerBI_DatasourceInventory.csv'
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
        8 { Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_DatasourceInventory' }
    }
}
