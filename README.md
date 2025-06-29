#  Excel‑to‑Hyper Converter: Solar Energy Example

This project demonstrates how to convert an Excel file (`RWFD_Solar_Energy.xlsx`) into Tableau `.hyper` extract files using Python. It’s designed for data analysts—especially those familiar with Tableau—who are learning or revisiting Python.

You’ll explore two methods:
- **Method 1 – Pantab Express**: A quick, one-liner conversion using the Pantab library.
- **Method 2 – Hyper API Adventure**: A more hands-on approach using Tableau’s official Hyper API for full control over schema and data insertion.

---

##  Requirements

Install the required Python packages:

```bash
pip install pandas pantab tableauhyperapi
```

## Project Structure
project-root/
│
├── src/
│   ├── method1_pantab.py       # Pantab method (quick and simple)
│   └── method2_hyperapi.py     # Hyper API method (more hands-on control)
│
├── data/                       # Excel source file
├── output/                     # Location for created .hyper extracts
├── video/                      # Demo videos
├── .gitignore
└── README.md

## Method 1 – Pantab Express 
This method uses the Pantab library to convert a DataFrame to a .hyper file.

```python
import pandas as pd
import pantab
import time; start=time.time()
```

## Read Excel sheet into DataFrame
```python
df = pd.read_excel(
    r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\data\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)
```

## Define output path for .hyper file
```python
output_hyper_path = r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\output\RWFD_Solar_Energy_method1.hyper"
```

## Write DataFrame to .hyper extract
```python
pantab.frame_to_hyper(df, output_hyper_path, table="Extract")

print(f"Hyper file created at: {output_hyper_path}")
print(f"Done in {time.time() - start: .2f} seconds")
```
## Run it with
```bash
python src/method2_pantab.py
```


## Method 2 – Hyper API Adventure
This method gives you full control over the schema and data insertion process using Tableau’s Hyper API.

```python
import pandas as pd
import time
from tableauhyperapi import (
    HyperProcess, Connection, TableDefinition, TableName,
    Inserter, SqlType, CreateMode, Telemetry, Nullability
)

start = time.time()
```

## Load Excel data into a DataFrame
```python
df = pd.read_excel(
    r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\data\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)
```

## Define output path for .hyper file
```python
output_hyper_path = r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\output\RWFD_Solar_Energy_method2.hyper"
```

## Define schema based on data types
```python
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
```

# Create and write to Hyper file
```python
with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
    with Connection(endpoint=hyper.endpoint, database=output_hyper_path, create_mode=CreateMode.CREATE_AND_REPLACE) as connection:
        connection.catalog.create_table(table_def)
        with Inserter(connection, table_def) as inserter:
            inserter.add_rows(df.itertuples(index=False, name=None))
            inserter.execute()

print(f"Hyper file created at: {output_hyper_path}")
print(f"Done in {time.time() - start:.2f} seconds")
```

## Run it with:
```bash
python src/method2_tableau_hyper_api.py
```

## Demo
https://youtu.be/6EDEYxyzM_c

## What to Expect
- **Two .hyper files generated (one for each method).
- **You should be able to open the Tableau Extracts using Tableau Desktop to verify the data.
- **You should also receive an output that shows how long the entire process took in order to compare.


