import pandas as pd
import pantab
import time; start=time.time()

# Define input Excel file and sheet
excel_path = r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\data\RWFD_Solar_Energy.xlsx"
sheet_name = "Actuals"

# Read Excel sheet into DataFrame
df = pd.read_excel(excel_path, sheet_name=sheet_name)

# Define output path for .hyper file
output_hyper_path = r"C:\Users\jvije\Documents\OneDrive\DataDevQuest\DDQ2025-05_Excel_to_Hyper_File\excel-to-hyper-converter-main\output\RWFD_Solar_Energy_Method1_pantab.hyper"

# Write DataFrame to .hyper extract
pantab.frame_to_hyper(df, output_hyper_path, table="Extract")

# Print completion message
print(f"Hyper file created at: {output_hyper_path}")
print(f"Done in {time.time() - start: .2f} seconds")
