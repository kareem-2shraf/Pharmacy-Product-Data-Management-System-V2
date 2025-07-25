import sys
import pandas as pd

if len(sys.argv) != 5:
    print("Usage: python sync_app_sheet.py app_sheet.csv stock_sheet.csv price_sheet.csv output.csv")
    sys.exit(1)

app_csv, stock_csv, price_csv, output_csv = sys.argv[1:5]

# Load sheets
df_app = pd.read_csv(app_csv, dtype=str, keep_default_na=False)
df_stock = pd.read_csv(stock_csv, dtype=str, keep_default_na=False)
df_price = pd.read_csv(price_csv, dtype=str, keep_default_na=False)

# Define columns (zero-based)
app_sku_col = 8
app_price_col = 9
app_stock_col = 27
stock_code_col = 0
stock_total_col = 21
price_code_col = 14
price_value_col = 8

# Build lookup dictionaries
stock_lookup = {str(row[stock_code_col]).strip(): str(row[stock_total_col]).strip()
                for _, row in df_stock.iterrows() if len(row) > max(stock_code_col, stock_total_col)}
price_lookup = {str(row[price_code_col]).strip(): str(row[price_value_col]).strip()
                for _, row in df_price.iterrows() if len(row) > max(price_code_col, price_value_col)}

# Update App Sheet
def update_row(row):
    if len(row) > max(app_sku_col, app_price_col, app_stock_col):
        sku = str(row[app_sku_col]).strip()
        # Update price
        if sku in price_lookup:
            row[app_price_col] = price_lookup[sku]
        # Update stock
        if sku in stock_lookup:
            row[app_stock_col] = stock_lookup[sku]
    return row

df_app = df_app.apply(update_row, axis=1)

df_app.to_csv(output_csv, index=False, encoding='utf-8')
print(f"✓ App Sheet synchronized and saved to {output_csv}") 