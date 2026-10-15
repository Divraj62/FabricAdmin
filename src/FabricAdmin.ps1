# Requires MicrosoftPowerBIMgmt.Profile and MicrosoftPowerBIMgmt.Workspaces
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile) -or
    -not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Workspaces)) {
    throw "Power BI module missing. Install it with: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
}

Import-Module MicrosoftPowerBIMgmt.Profile -WarningAction SilentlyContinue
Import-Module MicrosoftPowerBIMgmt.Workspaces -WarningAction SilentlyContinue

$moduleNames = @(
    'Common',
    'Workspaces',
    'Datasources',
    'RefreshMonitoring',
    'Gateways',
    'Security',
    'Activity',
    'ImpactAnalysis'
)
foreach ($moduleName in $moduleNames) {
    $modulePath = Join-Path $PSScriptRoot ("modules/{0}.psm1" -f $moduleName)
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        throw "FabricAdmin module missing: $modulePath"
    }
    . ([scriptblock]::Create((Get-Content -Raw -LiteralPath $modulePath)))
}

$script:OutputFolder = Join-Path (Split-Path -Parent $PSScriptRoot) 'exports'
$script:LoggingEnabled = $false
if (-not (Test-Path -LiteralPath $script:OutputFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
}

try {
    Connect-PowerBIServiceAccount -WarningAction SilentlyContinue

    Write-Host "Loading Power BI workspace inventory..."
    Initialize-WorkspaceInventory

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
