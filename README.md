# FabricAdmin

PowerShell console for read-only Power BI and Fabric governance inventory, monitoring, security review, and impact analysis.

## Run

Install the `MicrosoftPowerBIMgmt` PowerShell package and start the console:

```powershell
Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser
pwsh -NoProfile -File ./src/FabricAdmin.ps1
```

The console signs in interactively and retrieves organization-scoped inventory. Use an account with the required Power BI/Fabric administrator permissions.

## Layout

- `src/FabricAdmin.ps1` loads the modules, authenticates, initializes inventory, and dispatches the console menu.
- `src/modules/` holds the workspace, refresh, gateway, datasource, security, activity, impact, and shared helper functions.
- `exports/` is the default destination for CSV output. Five core exports use stable names: `PowerBI_Workspaces.csv`, `PowerBI_RefreshHealth.csv`, `PowerBI_GatewayInventory.csv`, `PowerBI_DatasourceInventory.csv`, and `PowerBI_Permissions.csv`.
- `notebooks/FabricAdmin_Ingestion.py` merges the workspace export into the attached Fabric Lakehouse table `workspaceinventory`.
- `docs/` describes the lakehouse flow, refresh-health columns, and a notebook query example.

Power BI exports can contain workspace identifiers, user details, and operational metadata. Generated exports remain ignored by Git; do not commit tenant data.

## Fabric Ingestion

Copy `exports/PowerBI_Workspaces.csv` to the attached Lakehouse under `Files/fabric-admin/PowerBI_Workspaces.csv`, then run `notebooks/FabricAdmin_Ingestion.py` in a Fabric notebook attached to that Lakehouse. The notebook performs an idempotent merge by workspace `Id`; it does not delete rows that are absent from a later snapshot.

See [Lakehouse architecture](docs/LakehouseArchitecture.md) and [refresh-health schema](docs/RefreshHealthSchema.md).

## Validation

The GitHub Actions workflow parses the PowerShell files and compiles the Python examples. It does not sign in, access tenant data, or deploy resources.
