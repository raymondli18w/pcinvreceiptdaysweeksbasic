import os
import pandas as pd
from datetime import date
from pathlib import Path

# Configuration
folder_path = r"C:\Users\RaymondLi\OneDrive - 18wheels.ca\downloads may 30 2023\test6\dec 2025 all\dec 2025 1\receipt date counter"
today = date(2026, 1, 12)  # ⚠️ UPDATE THIS TO ACTUAL "TODAY" OR USE date.today()
start_of_month = today.replace(day=1)

# Helper: Check if filename matches "piece" or "report" (case-insensitive)
def is_target_file(file_path: Path) -> bool:
    stem_lower = file_path.stem.lower()
    return "piece" in stem_lower or "report" in stem_lower

# Find all .xlsx files matching criteria
xlsx_files = [
    f for f in Path(folder_path).glob("*.xlsx")
    if is_target_file(f)
]

if not xlsx_files:
    raise FileNotFoundError("No .xlsx files found with 'piece' or 'report' in the name.")

# Get the latest (most recently modified)
latest_file = max(xlsx_files, key=os.path.getmtime)
print(f"Processing: {latest_file.name}")

# Load data
df = pd.read_excel(latest_file, header=0)
df.columns = df.columns.astype(str).str.strip()

# --- Group column detection ---
group_col = None
if 'Group ID' in df.columns:
    group_col = 'Group ID'
elif 'LP' in df.columns:
    group_col = 'LP'
else:
    available = ", ".join(f"'{col}'" for col in df.columns)
    raise KeyError(f"Neither 'Group ID' nor 'LP' found. Available columns: {available}")

print(f"Using '{group_col}' as grouping column.")
df = df.dropna(subset=[group_col]).copy()

# --- CRITICAL: Parse Receipt Date as MM/DD/YYYY ---
if 'Receipt Date' not in df.columns:
    available = ", ".join(f"'{col}'" for col in df.columns)
    raise KeyError(f"'Receipt Date' not found. Available columns: {available}")

df['Receipt Date'] = pd.to_datetime(
    df['Receipt Date'],
    errors='coerce',
    dayfirst=False,              # Ensures 1/9/2026 = Jan 9, not Sep 1
    infer_datetime_format=True
).dt.date

df = df.dropna(subset=['Receipt Date']).copy()

# --- Numeric columns to sum ---
potential_numeric_cols = [
    'Count Qty On Hand', 'Net Weight On Hand', 'Alt 1 Qty On Hand',
    'Alt 2 Qty On Hand', 'Grs Weight On Hand', 'Count Qty Committed',
    'Count Qty Uncommitted', 'Count Qty On Hold'
]
numeric_cols = [col for col in potential_numeric_cols if col in df.columns]
other_cols = [col for col in df.columns if col not in numeric_cols and col != group_col]

# --- Group and aggregate ---
agg_dict = {col: 'sum' for col in numeric_cols}
agg_dict.update({col: 'first' for col in other_cols})
summary = df.groupby(group_col, sort=False).agg(agg_dict).reset_index()
summary.rename(columns={group_col: 'Group ID'}, inplace=True)

# --- Calculate aging ---
summary['Days Since Receipt Date'] = summary['Receipt Date'].apply(
    lambda d: (today - d).days
)
summary['Weeks Since Receipt Date'] = (summary['Days Since Receipt Date'] / 7).round(2)

def weeks_from_start_month(receipt_date):
    if receipt_date.year == today.year and receipt_date.month == today.month:
        return None
    return round((today - start_of_month).days / 7, 2)

summary['Weeks from Start of Month to Today'] = summary['Receipt Date'].apply(weeks_from_start_month)

# --- Save ---
output_file = os.path.join(folder_path, "aggregated_receipt_report_final.xlsx")
summary.to_excel(output_file, index=False)

print(f"✅ Success! Output saved to:\n{output_file}")
print(f"📊 Processed {len(summary)} unique groups.")