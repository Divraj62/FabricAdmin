function Initialize-WorkspaceInventory {
    $workspaces = @(Get-PowerBIWorkspace -Scope Organization -All -Include All -WarningAction SilentlyContinue)
    $deletedWorkspaces = @(Get-PowerBIWorkspace -Scope Organization -Deleted -All -WarningAction SilentlyContinue)
    $orphanedWorkspaces = @(Get-PowerBIWorkspace -Scope Organization -Orphaned -All -WarningAction SilentlyContinue)

    $deletedOrOrphaned = @($deletedWorkspaces + $orphanedWorkspaces | Sort-Object Id -Unique)
    $publicWorkspaces = @($workspaces | Where-Object { $_.Type -in @('Workspace', 'Group') })
    $personalWorkspaces = @($workspaces | Where-Object { $_.Type -eq 'PersonalGroup' })
    $deletedOrOrphanedPublic = @($deletedOrOrphaned | Where-Object { $_.Type -in @('Workspace', 'Group') })
    $deletedOrOrphanedPersonal = @($deletedOrOrphaned | Where-Object { $_.Type -eq 'PersonalGroup' })

    $script:Workspaces = $workspaces
    $script:PublicWorkspaces = $publicWorkspaces
    $script:PersonalWorkspaces = $personalWorkspaces
    $script:DeletedWorkspaces = $deletedWorkspaces
    $script:OrphanedWorkspaces = $orphanedWorkspaces
    $script:deletedOrOrphanedPublic = $deletedOrOrphanedPublic
    $script:deletedOrOrphanedPersonal = $deletedOrOrphanedPersonal
    $script:Capacities = @()
    $script:ScanWorkspaces = $null
    $script:RefreshInventory = $null
    $script:GatewayInventory = $null
    $script:ActivityInventory = $null
    $script:ServerInventory = $null
    $script:ServerSummaryRows = @()
    $script:ServerLookupFailures = 0

    Write-Host "`nPOWER BI SERVICE ADMIN SUMMARY"
    Write-Host "--------------------------------"
    if ($workspaceSelection -eq '1') {
        Write-WorkspaceSummary -Title 'PUBLIC WORKSPACES' -CurrentWorkspaces $publicWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPublic
    }
    elseif ($workspaceSelection -eq '2') {
        Write-WorkspaceSummary -Title 'PERSONAL WORKSPACES' -CurrentWorkspaces $personalWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPersonal
    }
    else {
        Write-WorkspaceSummary -Title 'PUBLIC WORKSPACES' -CurrentWorkspaces $publicWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPublic
        Write-WorkspaceSummary -Title 'PERSONAL WORKSPACES' -CurrentWorkspaces $personalWorkspaces -DeletedOrOrphanedWorkspaces $deletedOrOrphanedPersonal
    }
}

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

function Get-WorkspaceCapacityName {
    param ([object]$Workspace)

    if ($Workspace.CapacityId) {
        $capacity = $script:Capacities | Where-Object { $_.Id -eq $Workspace.CapacityId } | Select-Object -First 1
        if ($capacity) { return $capacity.DisplayName }
        return [string]$Workspace.CapacityId
    }
    return 'Shared'
}

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
