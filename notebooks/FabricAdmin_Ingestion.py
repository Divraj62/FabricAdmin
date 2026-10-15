from delta.tables import DeltaTable
from pyspark.sql import SparkSession
from pyspark.sql import functions as F


spark = SparkSession.builder.getOrCreate()
WORKSPACE_CSV_PATH = "Files/fabric-admin/PowerBI_Workspaces.csv"
TARGET_TABLE = "workspaceinventory"

workspace_source = spark.read.option("header", "true").option("inferSchema", "false").csv(WORKSPACE_CSV_PATH)
if "Id" not in workspace_source.columns:
    raise ValueError("Workspace export must contain an Id column")

workspace_snapshot = (
    workspace_source
    .filter(F.col("Id").isNotNull())
    .dropDuplicates(["Id"])
)

if not workspace_snapshot.take(1):
    raise ValueError("Workspace export contains no rows with an Id")

if spark.catalog.tableExists(TARGET_TABLE):
    if "Id" not in spark.table(TARGET_TABLE).columns:
        raise ValueError(f"Target table {TARGET_TABLE} must contain an Id column")

    (
        DeltaTable.forName(spark, TARGET_TABLE)
        .alias("target")
        .merge(workspace_snapshot.alias("source"), "target.Id = source.Id")
        .whenMatchedUpdateAll()
        .whenNotMatchedInsertAll()
        .execute()
    )
else:
    workspace_snapshot.write.format("delta").mode("errorifexists").saveAsTable(TARGET_TABLE)

print(f"Merged workspace inventory into {TARGET_TABLE}")