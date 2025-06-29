import pandas as pd
import time
from tableauhyperapi import (
    HyperProcess, Connection, TableDefinition, TableName,
    Inserter, SqlType, CreateMode, Telemetry, Nullability
)

# Start timing
start = time.time()

# Load Excel data
df = pd.read_excel(
    r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\data\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)

# Define output path
output_hyper_path = r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\output\RWFD_Solar_Energy_Method2_HyperAPI.hyper"

# Define table schema with basic type inference
columns = []
for col in df.columns:
    if pd.api.types.is_float_dtype(df[col]):
        sql_type = SqlType.double()
    elif pd.api.types.is_integer_dtype(df[col]):
        sql_type = SqlType.big_int()
    elif pd.api.types.is_datetime64_any_dtype(df[col]):
        sql_type = SqlType.date()
    else:
        sql_type = SqlType.text()
    columns.append(TableDefinition.Column(name=col, type=sql_type, nullability=Nullability.NULLABLE))

table_def = TableDefinition(table_name=TableName("Extract"), columns=columns)

# Create Hyper file and insert data
with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
    with Connection(endpoint=hyper.endpoint, database=output_hyper_path, create_mode=CreateMode.CREATE_AND_REPLACE) as connection:
        connection.catalog.create_table(table_def)
        with Inserter(connection, table_def) as inserter:
            inserter.add_rows(df.itertuples(index=False, name=None))
            inserter.execute()

# Done
print(f"Hyper file created at: {output_hyper_path}")
print(f"Done in {time.time() - start:.2f} seconds")

