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
