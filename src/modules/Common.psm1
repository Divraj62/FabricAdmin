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

    $stableNames = @(
        'PowerBI_Workspaces',
        'PowerBI_RefreshHealth',
        'PowerBI_GatewayInventory',
        'PowerBI_DatasourceInventory',
        'PowerBI_Permissions'
    )
    if ($BaseName -in $stableNames) {
        $defaultPath = Join-Path $script:OutputFolder ("{0}.csv" -f $BaseName)
    }
    else {
        $defaultPath = Join-Path $script:OutputFolder ("{0}_{1}.csv" -f $BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
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

function Show-ExportSection {
    $choice = Read-ConsoleMenu -Title 'EXPORT & REPORTING' -Options @('Workspace Inventory', 'Report Inventory', 'Semantic Model Inventory', 'Server Inventory', 'Database Inventory', 'Datasource Inventory', 'Query Inventory', 'Refresh Failures', 'Permissions', 'Activity', 'Complete Tenant Inventory', 'Back')
    switch ($choice) {
        1 { Export-ConsoleRows -Rows @($script:Workspaces | Select-Object Name, Id, Type, State, CapacityId) -BaseName 'PowerBI_Workspaces' }
        2 { Export-ConsoleRows -Rows @(Get-ConsoleContent -WorkspaceSet $script:Workspaces).Reports -BaseName 'PowerBI_Reports' }
        3 { Export-ConsoleRows -Rows @(Get-ConsoleContent -WorkspaceSet $script:Workspaces).Models -BaseName 'PowerBI_SemanticModels' }
        4 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_Servers' }
        5 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows @($script:ServerInventory | Select-Object Server, Database, DatasourceType | Sort-Object Server, Database, DatasourceType -Unique) -BaseName 'PowerBI_Databases' }
        6 { if ($null -eq $script:ServerInventory) { $r = Get-ServerInventory -Workspaces $script:Workspaces; $script:ServerInventory = @($r.Items) }; Export-ConsoleRows -Rows $script:ServerInventory -BaseName 'PowerBI_DatasourceInventory' }
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
