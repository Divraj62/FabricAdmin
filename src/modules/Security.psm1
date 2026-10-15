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
