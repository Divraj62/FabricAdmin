from pyspark.sql import SparkSession
from pyspark.sql import functions as F

spark = SparkSession.builder.getOrCreate()
workspace_inventory = spark.table("workspaceinventory")

(
    workspace_inventory.select("Name", "Id", "Type", "State", "CapacityId")
    .where(F.col("Id").isNotNull())
    .orderBy("Name")
    .limit(25)
    .show(truncate=False)
)