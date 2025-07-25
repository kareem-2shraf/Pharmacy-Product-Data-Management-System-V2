# Pharmacy Product Data Management System (Version 2)

## Description

A user-friendly system for managing pharmacy product data (stock, price, and app sheets). This version features a new menu flow, improved update reliability, and Python integration for efficient file conversion and sorting.

---

## Features

- Main menu with three choices:
  1. Process Orders (update sheets)
  2. Other Operations (advanced menu)
  3. Exit
- Automatic conversion of HTML/Excel files to CSV using Python
- Reliable update functions to prevent missing or failed data
- Search, update, add, synchronize, sort, and export product data
- Error handling and user prompts for failed operations

---

## How It Works

1. **Startup:**
   - Program displays a main menu:
     - 1: Process Orders (for updating sheets)
     - 2: Other Operations (shows more options)
     - 3: Exit
2. **File Input:**
   - User enters paths for stock, price, and app sheets (HTML, Excel, or CSV)
   - Program converts files to CSV if needed
3. **Menu Navigation:**
   - Choose to process orders or access other operations
   - Perform actions like search, update, add, sync, sort, export, and more
4. **Data Integrity:**
   - All updates are checked for success; user is notified of any issues

---

## Technologies Used

- C++17 (main logic, menu, file handling)
- Python 3 (file conversion and sorting)
- Windows API (console encoding)
- STL (vectors, strings, algorithms)

---

## Getting Started

### Prerequisites
- Windows 10 or later
- C++17 compiler
- Python 3.x
- Python packages:
  ```
  pip install pandas openpyxl
  ```

### Files
- `main.cpp` — Main program and menu
- `file_handler.cpp/.h` — File I/O
- `operations.cpp/.h` — Data operations
- `html_to_csv_converter.py` — Converts HTML/Excel to CSV
- `sort_csv_by_column.py` — Sorts CSV files
- `log.txt` — Operation logs

### Usage
1. Build the project (compile C++ files)
2. Run the executable
3. Enter file paths when prompted
4. Use the menu to process orders or access other operations
5. Python scripts are called automatically for conversion/sorting

---

## Improvements in Version 2
- Redesigned menu for easier navigation
- More reliable update and sync functions
- Python script replaces C++ for file conversion (faster, more robust)
- Better error handling and user feedback

---

## Troubleshooting
- If you see missing Python module errors, run:
  ```
  pip install pandas openpyxl
  ```
- Make sure file paths are correct
- Add Python to your PATH if not found

---

## 📄 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for terms.
