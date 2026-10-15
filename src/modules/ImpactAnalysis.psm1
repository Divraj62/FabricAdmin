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
