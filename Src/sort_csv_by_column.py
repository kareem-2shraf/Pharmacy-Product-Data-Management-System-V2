import sys
import pandas as pd
import re

if len(sys.argv) != 4:
    print("Usage: python sort_csv_by_column.py input.csv output.csv column_name")
    sys.exit(1)

input_csv, output_csv, column_name = sys.argv[1], sys.argv[2], sys.argv[3]

df = pd.read_csv(input_csv, dtype=str, keep_default_na=False)
if column_name not in df.columns:
    print(f"Column '{column_name}' not found in {input_csv}")
    sys.exit(1)

# Create a helper column for numeric sorting
_df_numeric = pd.to_numeric(df[column_name], errors='coerce')
df['_numeric_sort'] = _df_numeric

df_sorted = df.sort_values(by=['_numeric_sort', column_name], na_position='last', kind='mergesort')
df_sorted = df_sorted.drop(columns=['_numeric_sort'])

# Remove rows where the sku column contains two or more separate numbers (with any characters between)
number_pattern = re.compile(r'(\d+)')
def has_multiple_numbers(s):
    if pd.isna(s):
        return False
    return len(number_pattern.findall(str(s))) >= 2
mask = df_sorted[column_name].apply(has_multiple_numbers)
df_sorted = df_sorted[~mask]

# Remove rows where the sku (digits only) is longer than 6 digits

def is_too_long_number(s):
    if pd.isna(s):
        return False
    digits = ''.join(filter(str.isdigit, str(s)))
    return len(digits) > 6
mask_long = df_sorted[column_name].apply(is_too_long_number)
df_sorted = df_sorted[~mask_long]

df_sorted.to_csv(output_csv, index=False, encoding='utf-8')
print(f"Sorted {input_csv} by '{column_name}' numerically (Excel/WPS style), removed rows with two or more separate numbers in SKU, and removed numbers longer than 6 digits. Saved to {output_csv}")
