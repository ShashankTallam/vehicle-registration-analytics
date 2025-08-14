import pandas as pd
import glob
import os
import re

RAW_DATA_FOLDER = r"C:\Users\shash\OneDrive\Desktop\free\raw data"
FILE_PATTERN = "*.xlsx"

all_files = glob.glob(os.path.join(RAW_DATA_FOLDER, FILE_PATTERN))
df_list = []

def detect_vehicle_type(fname):
    fname = fname.upper()
    if "2W" in fname:
        return "2W"
    elif "3W" in fname:
        return "3W"
    elif "4W" in fname:
        return "4W"
    else:
        return "Unknown"

def detect_year(fname):
    match = re.search(r'20\d{2}', fname)
    if match:
        return int(match.group(0))
    else:
        return None

# Detect month from filename (Jan, Feb, etc.)
def detect_month(fname):
    match = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', fname, re.IGNORECASE)
    if match:
        # Convert month name to number
        month_str = match.group(0).capitalize()
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        return months.index(month_str) + 1
    else:
        return None

for file in all_files:
    temp_df = pd.read_excel(file)

    # Add vehicle type, year, and month
    temp_df['vehicle_type'] = detect_vehicle_type(file)
    temp_df['year'] = detect_year(file)
    temp_df['month'] = detect_month(file)
    df_list.append(temp_df)

# Combine all dataframes
df = pd.concat(df_list, ignore_index=True)

# === Reset S No as continuous numbers ===
df['S No'] = range(1, len(df) + 1)

print(f"✅ Combined {len(df)} rows from {len(all_files)} files.")
print(df[['S No', 'vehicle_type', 'year', 'month']].head(15))

# Save merged dataset
output_path = r"C:\Users\shash\OneDrive\Desktop\free\vehicle_registrations_combined.xlsx"
df.to_excel(output_path, index=False)
print(f"✅ Combined file saved at: {output_path}")