#!/usr/bin/env python3
"""
File to CSV Converter for Pharmacy Product Data Management System
Converts HTML table files and Excel (XLSX) files to CSV format
"""

import os
import sys
from pathlib import Path
import re
import argparse

# Try to import required libraries with helpful error messages
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("✗ pandas library not found. Please install it with: pip install pandas")
    print("For Excel support, also install: pip install openpyxl")

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("✗ openpyxl library not found. Excel files will not be supported.")
    print("Install with: pip install openpyxl")

def clean_html_content(html_content):
    """Clean HTML content by removing extra whitespace and formatting"""
    # Remove extra whitespace
    html_content = re.sub(r'\s+', ' ', html_content)
    # Remove HTML comments
    html_content = re.sub(r'<!--.*?-->', '', html_content)
    return html_content

def extract_table_from_html(file_path):
    """Extract table data from HTML file"""
    if not PANDAS_AVAILABLE:
        print(f"✗ Cannot convert HTML file: pandas not available")
        return None

    print(f"Converting HTML file: {file_path}")

    try:
        # Read the HTML file
        with open(file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Clean the HTML content
        html_content = clean_html_content(html_content)

        # Read all tables from HTML
        tables = pd.read_html(html_content, header=0)

        if not tables:
            print(f"✗ No tables found in {file_path}")
            return None

        print(f"✓ Found {len(tables)} table(s) in {file_path}")

        # Use the first table (most common case)
        df = tables[0]

        print(f"✓ Extracted table with {len(df)} rows and {len(df.columns)} columns")
        return df

    except Exception as e:
        print(f"✗ Error converting {file_path}: {e}")
        return None

def extract_table_from_xlsx(file_path):
    """Extract table data from XLSX file"""
    if not PANDAS_AVAILABLE:
        print(f"✗ Cannot convert Excel file: pandas not available")
        return None

    if not OPENPYXL_AVAILABLE:
        print(f"✗ Cannot convert Excel file: openpyxl not available")
        return None

    print(f"Converting XLSX file: {file_path}")

    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=0)

        if df.empty:
            print(f"✗ No data found in {file_path}")
            return None

        print(f"✓ Extracted table with {len(df)} rows and {len(df.columns)} columns")
        return df

    except Exception as e:
        print(f"✗ Error converting {file_path}: {e}")
        return None

def convert_file_to_csv(input_path, output_path):
    """Convert HTML or XLSX file to CSV"""
    file_extension = Path(input_path).suffix.lower()

    if file_extension in ['.html', '.htm']:
        df = extract_table_from_html(input_path)
    elif file_extension in ['.xlsx', '.xls']:
        df = extract_table_from_xlsx(input_path)
    else:
        print(f"✗ Unsupported file format: {file_extension}")
        print("Supported formats: .html, .htm, .xlsx, .xls")
        return False

    if df is not None:
        try:
            # Save as CSV
            df.to_csv(output_path, index=False, encoding='utf-8')
            print(f"✓ Successfully converted to: {output_path}")
            return True
        except Exception as e:
            print(f"✗ Error saving CSV: {e}")
            return False
    return False

def csv_to_xlsx(input_csv, output_xlsx):
    try:
        df = pd.read_csv(input_csv, dtype=str, keep_default_na=False)
        # Convert every cell to a number if possible
        def try_numeric(val):
            try:
                # Try integer first, then float
                if val.strip() == '':
                    return val
                if '.' in val:
                    f = float(val)
                    if f.is_integer():
                        return int(f)
                    return f
                return int(val)
            except Exception:
                return val
        df = df.applymap(try_numeric)
        df.to_excel(output_xlsx, index=False, engine='openpyxl')
        print(f"✓ Successfully converted {input_csv} to {output_xlsx} (all numeric cells as numbers)")
        return True
    except Exception as e:
        print(f"✗ Error converting CSV to XLSX: {e}")
        return False

def main():
    """Main function to convert HTML/XLSX files to CSV or CSV to XLSX"""
    parser = argparse.ArgumentParser(description='Convert HTML/Excel files to CSV or CSV to XLSX')
    parser.add_argument('--single-file', nargs=2, metavar=('INPUT', 'OUTPUT'),
                       help='Convert a single file to CSV')
    parser.add_argument('--csv-to-xlsx', nargs=2, metavar=('INPUT_CSV', 'OUTPUT_XLSX'),
                       help='Convert a CSV file to XLSX')

    args = parser.parse_args()

    # CSV to XLSX conversion mode
    if args.csv_to_xlsx:
        input_csv, output_xlsx = args.csv_to_xlsx
        if not os.path.exists(input_csv):
            print(f"✗ Input CSV file not found: {input_csv}")
            sys.exit(1)
        success = csv_to_xlsx(input_csv, output_xlsx)
        sys.exit(0 if success else 1)

    # Single file conversion mode (called from C++)
    if args.single_file:
        input_path, output_path = args.single_file
        if not os.path.exists(input_path):
            print(f"✗ Input file not found: {input_path}")
            sys.exit(1)

        success = convert_file_to_csv(input_path, output_path)
        sys.exit(0 if success else 1)

    # Interactive mode (default)
    print("=" * 60)
    print("File to CSV Converter for Pharmacy Data Management System")
    print("Supports: HTML (.html, .htm) and Excel (.xlsx, .xls) files")
    print("=" * 60)

    # Check dependencies
    if not PANDAS_AVAILABLE:
        print("✗ Required dependency 'pandas' not found.")
        print("Please install with: pip install pandas")
        return

    # Get file paths from user
    stock_file = input("Enter path for Stock Sheet (HTML/XLSX): ").strip()
    price_file = input("Enter path for Price Sheet (HTML/XLSX): ").strip()
    app_file = input("Enter path for App Sheet (HTML/XLSX): ").strip()

    # Validate input files
    files_to_check = [
        (stock_file, "Stock Sheet"),
        (price_file, "Price Sheet"),
        (app_file, "App Sheet")
    ]

    for file_path, file_name in files_to_check:
        if not os.path.exists(file_path):
            print(f"✗ {file_name} not found: {file_path}")
            return
        file_extension = Path(file_path).suffix.lower()
        if file_extension not in ['.html', '.htm', '.xlsx', '.xls']:
            print(f"✗ {file_name} is not a supported file format: {file_path}")
            print("Supported formats: .html, .htm, .xlsx, .xls")
            return

    print("\nStarting file to CSV conversion...")
    print("-" * 40)

    # Convert each file
    success_count = 0

    # Stock sheet
    stock_csv = str(Path(stock_file).with_suffix('.csv'))
    if convert_file_to_csv(stock_file, stock_csv):
        success_count += 1

    # Price sheet
    price_csv = str(Path(price_file).with_suffix('.csv'))
    if convert_file_to_csv(price_file, price_csv):
        success_count += 1

    # App sheet
    app_csv = str(Path(app_file).with_suffix('.csv'))
    if convert_file_to_csv(app_file, app_csv):
        success_count += 1

    print("-" * 40)
    if success_count == 3:
        print("✓ All files converted successfully!")
        print(f"\nConverted files:")
        print(f"  Stock: {stock_csv}")
        print(f"  Price: {price_csv}")
        print(f"  App: {app_csv}")
        print(f"\nYou can now use these CSV files with your C++ program!")
    else:
        print(f"⚠ {success_count}/3 files converted successfully")
    print("=" * 60)

if __name__ == "__main__":
    main()