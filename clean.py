import pandas as pd

# ====== CONFIG ======
INPUT_FILE = r"C:\Users\shash\OneDrive\Desktop\free\vehicle_registrations_combined.xlsx"
OUTPUT_FILE = r"C:\Users\shash\OneDrive\Desktop\free\vehicle_registrations_cleaned.xlsx"

# ====== 1. LOAD DATA ======
df = pd.read_excel(INPUT_FILE)
print("✅ File loaded")
print("Columns before cleaning:", df.columns.tolist())

# ====== 2. STRIP SPACES ======
df.columns = df.columns.str.strip()

# ====== 3. FIX 'S No' COLUMN ======
if 'S No' not in df.columns:
    first_col = df.columns[0]
    print(f"Renaming '{first_col}' to 'S No'")
    df.rename(columns={first_col: 'S No'}, inplace=True)

# Always clean below here (not inside the if block)
df = df[pd.to_numeric(df['S No'], errors='coerce').notnull()]
df['S No'] = pd.to_numeric(df['S No'], errors='coerce').astype(int)

# ====== 4. STANDARDIZE MANUFACTURER ======
for col in df.columns:
    if "maker" in col.lower() or "manufactur" in col.lower():
        df.rename(columns={col: 'manufacturer'}, inplace=True)
        break

# ====== 5. FIX REGISTRATIONS COLUMN ======
possible_reg_cols = [c for c in df.columns if 'reg' in c.lower()]
if not possible_reg_cols:
    possible_reg_cols = [c for c in df.columns if df[c].dtype in ['int64', 'float64'] and c not in ['year', 'S No', 'month']]
if possible_reg_cols and 'registrations' not in df.columns:
    df.rename(columns={possible_reg_cols[-1]: 'registrations'}, inplace=True)

if 'registrations' in df.columns:
    df['registrations'] = pd.to_numeric(df['registrations'], errors='coerce').fillna(0).astype(int)

# ====== 6. REMOVE BLANKS/ZEROS ======
if 'manufacturer' in df.columns:
    df['manufacturer'] = df['manufacturer'].astype(str).str.strip()
    df = df[df['manufacturer'] != '']

if 'registrations' in df.columns:
    df = df[df['registrations'] > 0]


# ====== 7. OPTIONAL: PIVOT TO MONTH-WISE ======
# Drop any existing S No column before pivoting
if 'S No' in df.columns:
    df = df.drop(columns=['S No'])

if 'month' in df.columns and 'registrations' in df.columns:
    # Remove duplicates so each manufacturer/year/vehicle_type/month is unique
    pivot_cols = ['manufacturer', 'year', 'vehicle_type']
    if 'vehicle_category' in df.columns:
        pivot_cols.append('vehicle_category')
    df_grouped = df.groupby(pivot_cols + ['month'], as_index=False)['registrations'].sum()
    df_pivot = df_grouped.pivot_table(index=pivot_cols, columns='month', values='registrations', fill_value=0)
    df_pivot.columns = [f"month_{int(col)}" for col in df_pivot.columns]
    df_pivot = df_pivot.reset_index()
    # Now set S No as continuous
    df_pivot['S No'] = range(1, len(df_pivot) + 1)
    df_final = df_pivot
else:
    df_final = df.copy()
    # If S No exists, reset it as continuous
    if 'S No' in df_final.columns:
        df_final['S No'] = range(1, len(df_final) + 1)

# ====== 8. SAVE CLEANED FILE ======
print("Saving file with shape:", df_final.shape)
df_final.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Cleaned file saved to: {OUTPUT_FILE}")
