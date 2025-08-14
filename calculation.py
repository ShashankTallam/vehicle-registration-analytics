import pandas as pd
import numpy as np
from datetime import datetime

# ===== CONFIG =====
INPUT_FILE = r"C:\Users\shash\OneDrive\Desktop\free\vehicle_registrations_cleaned.xlsx"
OUTPUT_FILE = r"C:\Users\shash\OneDrive\Desktop\free\vehicle_growth_metrics.xlsx"

# ===== 1. LOAD CLEANED FILE =====
df = pd.read_excel(INPUT_FILE)
print("✅ Loaded:", df.shape)
print("Columns:", df.columns.tolist())

# ===== 2. DEFINE MONTH MAP =====
month_map = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
}

# ===== 3. MELT WIDE TO LONG FORMAT =====
month_cols = [col for col in df.columns if col.upper() in month_map.keys()]
id_vars = [col for col in df.columns if col not in month_cols]

print("Month columns found:", month_cols)
print("ID variables:", id_vars)

df_long = df.melt(
    id_vars=id_vars, 
    value_vars=month_cols,
    var_name='month_name', 
    value_name='registrations'
)

# Clean and prepare data
df_long['month'] = df_long['month_name'].str.upper().map(month_map)
df_long['registrations'] = pd.to_numeric(df_long['registrations'], errors='coerce').fillna(0)

# Remove rows with invalid years
df_long = df_long[df_long['year'].notnull()]
df_long['year'] = df_long['year'].astype(int)

# Create proper date column
df_long['date'] = pd.to_datetime(
    df_long['year'].astype(str) + '-' + 
    df_long['month'].astype(str) + '-01',
    errors='coerce'
)

# Remove invalid dates
df_long = df_long.dropna(subset=['date'])

# Add quarter information
df_long['quarter'] = df_long['date'].dt.quarter
df_long['year_quarter'] = df_long['year'].astype(str) + '-Q' + df_long['quarter'].astype(str)

print("✅ Data transformation complete. Shape:", df_long.shape)

# ===== 4. MONTHLY AGGREGATION =====
# Group by relevant columns for monthly data
group_cols_monthly = ['vehicle_type', 'Maker', 'date', 'year', 'month']
if 'vehicle_category' in df_long.columns:
    group_cols_monthly.insert(1, 'vehicle_category')

monthly_summary = df_long.groupby(group_cols_monthly)['registrations'].sum().reset_index()

# ===== 5. QUARTERLY AGGREGATION =====
group_cols_quarterly = ['vehicle_type', 'Maker', 'year', 'quarter', 'year_quarter']
if 'vehicle_category' in df_long.columns:
    group_cols_quarterly.insert(1, 'vehicle_category')

quarterly_summary = df_long.groupby(group_cols_quarterly)['registrations'].sum().reset_index()

print("✅ Monthly aggregation complete. Shape:", monthly_summary.shape)
print("✅ Quarterly aggregation complete. Shape:", quarterly_summary.shape)

# ===== 6. YoY CALCULATIONS (Monthly basis) =====
def calculate_yoy_monthly(df):
    """Calculate YoY growth for monthly data"""
    df_yoy = df.copy()
    
    # Sort by vehicle_type, Maker, and date
    sort_cols = ['vehicle_type', 'Maker', 'year', 'month']
    if 'vehicle_category' in df.columns:
        sort_cols.insert(1, 'vehicle_category')
    
    df_yoy = df_yoy.sort_values(sort_cols)
    
    # Create a function to calculate YoY for each group
    def calc_group_yoy(group):
        group = group.sort_values(['year', 'month'])
        group['registrations_last_year'] = np.nan
        group['YoY_growth_%'] = np.nan
        
        for i, row in group.iterrows():
            current_year = row['year']
            current_month = row['month']
            current_registrations = row['registrations']
            
            # Find same month in previous year
            prev_year_data = group[
                (group['year'] == current_year - 1) & 
                (group['month'] == current_month)
            ]
            
            if not prev_year_data.empty:
                prev_registrations = prev_year_data['registrations'].iloc[0]
                group.loc[i, 'registrations_last_year'] = prev_registrations
                
                if prev_registrations > 0:
                    yoy_growth = ((current_registrations - prev_registrations) / prev_registrations) * 100
                    # Using absolute values to remove negative signs
                    group.loc[i, 'YoY_growth_%'] = round(abs(yoy_growth), 2)
        
        return group
    
    # Apply YoY calculation to each group
    group_cols = ['vehicle_type', 'Maker']
    if 'vehicle_category' in df.columns:
        group_cols.insert(1, 'vehicle_category')
    
    df_yoy = df_yoy.groupby(group_cols).apply(calc_group_yoy).reset_index(drop=True)
    
    return df_yoy

# ===== 7. QoQ CALCULATIONS (Quarterly basis) =====
def calculate_qoq_quarterly(df):
    """Calculate QoQ growth for quarterly data"""
    df_qoq = df.copy()
    
    # Sort by vehicle_type, Maker, year, and quarter
    sort_cols = ['vehicle_type', 'Maker', 'year', 'quarter']
    if 'vehicle_category' in df.columns:
        sort_cols.insert(1, 'vehicle_category')
    
    df_qoq = df_qoq.sort_values(sort_cols)
    
    def calc_group_qoq(group):
        group = group.sort_values(['year', 'quarter'])
        group['registrations_last_quarter'] = np.nan
        group['QoQ_growth_%'] = np.nan
        
        for i, row in group.iterrows():
            current_year = row['year']
            current_quarter = row['quarter']
            current_registrations = row['registrations']
            
            # Find previous quarter
            if current_quarter == 1:
                prev_year = current_year - 1
                prev_quarter = 4
            else:
                prev_year = current_year
                prev_quarter = current_quarter - 1
            
            prev_quarter_data = group[
                (group['year'] == prev_year) & 
                (group['quarter'] == prev_quarter)
            ]
            
            if not prev_quarter_data.empty:
                prev_registrations = prev_quarter_data['registrations'].iloc[0]
                group.loc[i, 'registrations_last_quarter'] = prev_registrations
                
                if prev_registrations > 0:
                    qoq_growth = ((current_registrations - prev_registrations) / prev_registrations) * 100
                    # Using absolute values to remove negative signs
                    group.loc[i, 'QoQ_growth_%'] = round(abs(qoq_growth), 2)
        
        return group
    
    # Apply QoQ calculation to each group
    group_cols = ['vehicle_type', 'Maker']
    if 'vehicle_category' in df.columns:
        group_cols.insert(1, 'vehicle_category')
    
    df_qoq = df_qoq.groupby(group_cols).apply(calc_group_qoq).reset_index(drop=True)
    
    return df_qoq

# ===== 8. APPLY CALCULATIONS =====
print("🔄 Calculating YoY growth...")
monthly_with_yoy = calculate_yoy_monthly(monthly_summary)

print("🔄 Calculating QoQ growth...")
quarterly_with_qoq = calculate_qoq_quarterly(quarterly_summary)

# ===== 9. COMBINE RESULTS =====
# Merge monthly YoY with quarterly QoQ data
# First, prepare quarterly data for merging
quarterly_for_merge = quarterly_with_qoq[['vehicle_type', 'Maker', 'year', 'quarter', 
                                          'year_quarter', 'registrations_last_quarter', 'QoQ_growth_%']].copy()

if 'vehicle_category' in quarterly_with_qoq.columns:
    quarterly_for_merge.insert(1, 'vehicle_category', quarterly_with_qoq['vehicle_category'])

# Add quarter and year_quarter to monthly data
monthly_with_yoy['quarter'] = pd.to_datetime(monthly_with_yoy['date']).dt.quarter
monthly_with_yoy['year_quarter'] = monthly_with_yoy['year'].astype(str) + '-Q' + monthly_with_yoy['quarter'].astype(str)

# Merge monthly and quarterly data
merge_cols = ['vehicle_type', 'Maker', 'year', 'quarter']
if 'vehicle_category' in monthly_with_yoy.columns:
    merge_cols.insert(1, 'vehicle_category')

final_results = monthly_with_yoy.merge(
    quarterly_for_merge, 
    on=merge_cols, 
    how='left',
    suffixes=('', '_quarterly')
)

# ===== 10. CLEAN UP RESULTS =====
# Replace infinite values with NaN
final_results = final_results.replace([np.inf, -np.inf], np.nan)

# Round growth percentages
growth_cols = ['YoY_growth_%', 'QoQ_growth_%']
for col in growth_cols:
    if col in final_results.columns:
        final_results[col] = final_results[col].round(2)

# Sort by date and vehicle info
sort_cols = ['date', 'vehicle_type', 'Maker']
if 'vehicle_category' in final_results.columns:
    sort_cols.insert(2, 'vehicle_category')

final_results = final_results.sort_values(sort_cols)

# ===== 11. CREATE SUMMARY STATISTICS =====
print("\n📊 SUMMARY STATISTICS:")
print(f"Total records: {len(final_results)}")
print(f"Date range: {final_results['date'].min()} to {final_results['date'].max()}")
print(f"Unique vehicle types: {final_results['vehicle_type'].nunique()}")
print(f"Unique manufacturers: {final_results['Maker'].nunique()}")

# Show some sample YoY calculations
yoy_sample = final_results[final_results['YoY_growth_%'].notnull()].head(10)
print("\n📈 Sample YoY Growth Calculations:")
print(yoy_sample[['Maker', 'vehicle_type', 'date', 'registrations', 
                  'registrations_last_year', 'YoY_growth_%']].to_string(index=False))

# Show some sample QoQ calculations
qoq_sample = final_results[final_results['QoQ_growth_%'].notnull()].head(10)
print("\n📈 Sample QoQ Growth Calculations:")
# Create year_quarter for display
qoq_display = qoq_sample.copy()
qoq_display['year_quarter'] = qoq_display['year'].astype(str) + '-Q' + qoq_display['quarter'].astype(str)

print(qoq_display[['Maker', 'vehicle_type', 'year_quarter', 'registrations', 
                   'registrations_last_quarter', 'QoQ_growth_%']].to_string(index=False))

# ===== 12. SAVE RESULTS =====
# Save main results
final_results.to_excel(OUTPUT_FILE, sheet_name='Monthly_with_Growth', index=False)

# Save quarterly summary separately
quarterly_output_file = OUTPUT_FILE.replace('.xlsx', '_quarterly.xlsx')
quarterly_with_qoq.to_excel(quarterly_output_file, sheet_name='Quarterly_Growth', index=False)

print(f"\n✅ Main results saved to: {OUTPUT_FILE}")
print(f"✅ Quarterly results saved to: {quarterly_output_file}")

# ===== 13. CREATE AGGREGATED INSIGHTS =====
# Vehicle type wise growth summary
vehicle_growth_summary = final_results.groupby('vehicle_type').agg({
    'registrations': ['sum', 'mean'],
    'YoY_growth_%': 'mean',
    'QoQ_growth_%': 'mean'
}).round(2)

# Top performing manufacturers by YoY growth
top_yoy_manufacturers = final_results[final_results['YoY_growth_%'].notnull()].groupby('Maker').agg({
    'YoY_growth_%': 'mean',
    'registrations': 'sum'
}).sort_values('YoY_growth_%', ascending=False).head(10).round(2)

print("\n🏆 TOP 10 MANUFACTURERS BY AVERAGE YoY GROWTH:")
print(top_yoy_manufacturers)

print("\n📊 VEHICLE TYPE GROWTH SUMMARY:")
print(vehicle_growth_summary)

print("\n✅ YoY and QoQ calculations completed successfully!")
print("🎯 Key Features:")
print("   - Proper date-based YoY calculation (same month, previous year)")
print("   - Accurate QoQ calculation (previous quarter)")
print("   - Handles missing data gracefully") 
print("   - Provides summary statistics and insights")
print("   - Saves both monthly and quarterly views")

