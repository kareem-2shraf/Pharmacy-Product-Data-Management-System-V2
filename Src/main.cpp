#include <iostream>
#include <string>
#include <vector>
#include <cstdlib>
#include <filesystem>
#include <algorithm>
#include <windows.h>
#include "file_handler.h"
#include "operations.h"
#include <chrono>
#include <thread>


//color codes
#define RESET   "\033[0m"
#define RED     "\033[31m"
#define GREEN   "\033[32m"
#define YELLOW  "\033[33m"
#define BLUE    "\033[34m"
#define MAGENTA "\033[35m"
#define CYAN    "\033[36m"
using namespace std;

// Function to strip quotes from file path
/*string stripQuotes(const string& path) {
    string result = path;
    // Remove leading and trailing quotes
    if (!result.empty() && (result.front() == '"' || result.front() == '\'')) {
        result = result.substr(1);
    }
    if (!result.empty() && (result.back() == '"' || result.back() == '\'')) {
        result = result.substr(0, result.length() - 1);
    }
    return result;
} */


// Function to check if file is HTML or Excel
bool isConvertibleFile(const string& filePath) {
    string lowerPath = filePath;
    transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::tolower);
    return lowerPath.find(".html") != string::npos ||
           lowerPath.find(".htm") != string::npos ||
           lowerPath.find(".xlsx") != string::npos ||
           lowerPath.find(".xls") != string::npos;
}

// Function to convert HTML/Excel to CSV using Python script
bool convertFileToCsv(const string& inputPath, const string& csvPath) {
    cout << "Attempting to convert: " << inputPath << " to " << csvPath << endl;

    // Try different Python commands
    vector<string> pythonCommands = {
        "python html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\"",
        "python3 html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\"",
        "py html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\"",
        "C:\\Python39\\python.exe html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\"",
        "C:\\Python310\\python.exe html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\"",
        "C:\\Python311\\python.exe html_to_csv_converter.py --single-file \"" + inputPath + "\" \"" + csvPath + "\""
    };

    for (const string& command : pythonCommands) {
        cout << "Trying command: " << command << endl;
        int result = system(command.c_str());
        cout << "Command result: " << result << endl;
        if (result == 0) {
            cout << "✓ Conversion successful!" << endl;
            return true;
        }
    }

    cout << "✗ All Python commands failed. Please ensure Python and required libraries are installed:" << endl;
    cout << "  pip install pandas openpyxl" << endl;
    return false;
}

void Delay(int t) {
    this_thread::sleep_for(chrono::milliseconds(t));
}

int main() {

    // Set console to UTF-8 for proper Arabic text display
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);


    cout << GREEN <<"******************************************************************" << RESET << endl;
    Delay(200);
    cout  << "\tWelcome to the ";
    Delay(500);
    cout << BLUE << "Pharmacy Product " << RESET ;
    Delay(500);
    cout << RED <<"Data Management " << RESET ;
    Delay(500);
    cout << "System" << RESET << endl;
    Delay(200);
    cout << GREEN <<"******************************************************************" << RESET << endl;
    cout << YELLOW << "Please enter the paths for the files" << RESET << endl;
    cout << BLUE << "*********************************************************" << RESET << endl;
    string stockPath, pricePath, appPath;
    cout << "Enter path for Stock Sheet: ";
    getline(cin, stockPath);
    //stockPath = stripQuotes(stockPath);
    cout << "Enter path for Price Sheet: ";
    getline(cin, pricePath);
    //stockPath = stripQuotes(pricePath);
    cout << "Enter path for App Sheet: ";
    getline(cin, appPath);
    //stockPath = stripQuotes(appPath);

    // Check if stock and price files are HTML/Excel and convert them
    if (isConvertibleFile(stockPath)) {
        cout << "Detected convertible file for Stock Sheet. Converting to CSV..." << endl;
        string stockCsvPath = stockPath.substr(0, stockPath.find_last_of('.')) + ".csv";
        if (convertFileToCsv(stockPath, stockCsvPath)) {
            cout << "✓ Stock Sheet converted to: " << stockCsvPath << endl;
            stockPath = stockCsvPath;
        } else {
            cout << "✗ Failed to convert Stock Sheet to CSV" << endl;
            cout << "Do you want to continue with the original file? (y/n): ";
            char choice;
            cin >> choice;
            if (choice != 'y' && choice != 'Y') {
                return 1;
            }
        }
    }

    if (isConvertibleFile(pricePath)) {
        cout << "Detected convertible file for Price Sheet. Converting to CSV..." << endl;
        string priceCsvPath = pricePath.substr(0, pricePath.find_last_of('.')) + ".csv";
        if (convertFileToCsv(pricePath, priceCsvPath)) {
            cout << "✓ Price Sheet converted to: " << priceCsvPath << endl;
            pricePath = priceCsvPath;
        } else {
            cout << "✗ Failed to convert Price Sheet to CSV" << endl;
            cout << "Do you want to continue with the original file? (y/n): ";
            char choice;
            cin >> choice;
            if (choice != 'y' && choice != 'Y') {
                return 1;
            }
        }
    }

    if (isConvertibleFile(appPath)) {
        cout << "Detected convertible file for App Sheet. Converting to CSV..." << endl;
        string appCsvPath = appPath.substr(0, appPath.find_last_of('.')) + ".csv";
        if (convertFileToCsv(appPath, appCsvPath)) {
            cout << "✓ App Sheet converted to: " << appCsvPath << endl;
            appPath = appCsvPath;
        } else {
            cout << "✗ Failed to convert App Sheet to CSV" << endl;
            cout << "Do you want to continue with the original file? (y/n): ";
            char choice;
            cin >> choice;
            if (choice != 'y' && choice != 'Y') {
                return 1;
            }
        }
    }


    SheetData stockSheet, priceSheet, appSheet;
    if (!readCSV(stockPath, stockSheet)) {
        cout << RED << "Failed to load Stock Sheet. Exiting.\n";
        return 1;
    }
    if (!readCSV(pricePath, priceSheet)) {
        cout << "Failed to load Price Sheet. Exiting.\n";
        return 1;
    }
    if (!readCSV(appPath, appSheet)) {
        cout << "Failed to load App Sheet. Exiting.\n" << RESET;
        return 1;
    }


    while (true) {
        cout << YELLOW << "\n--- Pharmacy Product Data Management System ---\n" << RESET;
        cout << "1. Search for a product by code\n";
        cout << "2. Update stock in Stock Sheet\n";
        cout << "3. Update price in Price Sheet\n";
        cout << "4. Add a new product\n";
        cout << "5. Synchronize and update App Sheet\n";
        cout << "6. Sort App Sheet by sku & Remove rows with multiple numbers in SKU in App Sheet\n";
        cout << "7. Export App Sheet\n";
        cout << "8. Save All Sheets and Exit\n";
        cout << "9. Convert App Sheet CSV to XLSX\n";
        cout << "10. Delete first 10 rows from Price Sheet\n";
        cout << "11. Unmerge columns I & J in Price Sheet\n";
        cout << "12. Move P column (الكود) to H column in Price Sheet\n";
        cout << "13. Delete repeated (السعر) column (index 9) from Price Sheet\n";
        cout << "14. Delete first 6 columns from Price Sheet\n";
        cout << "15. Set home nursing services stock to 9000000 in App Sheet\n";
        cout << "16. Convert all nan values to 0 in App Sheet\n";
        cout << "Select an option: ";
        int choice;
        cin >> choice;
        cin.ignore();
        switch (choice) {
            case 1: {
                cout << "Search in which sheet? (1=Stock, 2=Price, 3=App): ";
                int searchSheet;
                cin >> searchSheet;
                cin.ignore();
                string code;
                cout << "Enter product code: ";
                getline(cin, code);
                if (searchSheet == 1) {
                    searchProduct(code, stockSheet, SheetData(), SheetData());
                } else if (searchSheet == 2) {
                    searchProduct(code, SheetData(), priceSheet, SheetData());
                } else if (searchSheet == 3) {
                    searchProduct(code, SheetData(), SheetData(), appSheet);
                } else {
                    cout << "Invalid sheet choice.\n";
                }
                break;
            }
            case 2: {
                string code;
                int newStock;
                cout << "Enter product code: ";
                getline(cin, code);
                cout << "Enter new stock amount: ";
                cin >> newStock;
                cin.ignore();
                updateStock(stockSheet, code, newStock);
                break;
            }
            case 3: {
                string code;
                double newPrice;
                cout << "Enter product code: ";
                getline(cin, code);
                cout << "Enter new price: ";
                cin >> newPrice;
                cin.ignore();
                updatePrice(priceSheet, code, newPrice);
                break;
            }
            case 4: {
                cout << "Add to which sheet? (1=Stock, 2=Price, 3=App): ";
                int sheetChoice;
                cin >> sheetChoice;
                cin.ignore();
                vector<string> row;
                if (sheetChoice == 1) {
                    string code, total;
                    cout << "Enter code: "; getline(cin, code);
                    cout << "Enter total: "; getline(cin, total);
                    row.push_back(code); row.push_back(total);
                    addProduct(stockSheet, row);
                } else if (sheetChoice == 2) {
                    string code, price;
                    cout << "Enter code: "; getline(cin, code);
                    cout << "Enter price: "; getline(cin, price);
                    row.push_back(code); row.push_back(price);
                    addProduct(priceSheet, row);
                } else if (sheetChoice == 3) {
                    string sku, price, stock;
                    cout << "Enter sku: "; getline(cin, sku);
                    cout << "Enter price: "; getline(cin, price);
                    cout << "Enter stock: "; getline(cin, stock);
                    row.push_back(sku); row.push_back(price); row.push_back(stock);
                    addProduct(appSheet, row);
                } else {
                    cout << "Invalid sheet choice.\n";
                }
                break;
            }
            case 5:
                syncAppSheet(appSheet, stockSheet, priceSheet);
                break;
            case 6:{
                // Sort App Sheet in-place using Python (Excel/WPS style)
                string inOut = appPath;
                vector<string> pythonSortCmds = {
                    "python sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku",
                    "python3 sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku",
                    "py sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku",
                    "C:\\Python39\\python.exe sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku",
                    "C:\\Python310\\python.exe sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku",
                    "C:\\Python311\\python.exe sort_csv_by_column.py \"" + inOut + "\" \"" + inOut + "\" sku"
                };
                bool success = false;
                for (const string& cmd : pythonSortCmds) {
                    cout << "Trying command: " << cmd << endl;
                    int result = system(cmd.c_str());
                    if (result == 0 && readCSV(inOut, appSheet)) {
                        cout << "App Sheet sorted by SKU (ascending === numerically === with deleting The Offer Rows)...\n";
                        success = true;
                        break;
                    }
                }
                if (!success) {
                    cout << "Sorting failed or could not reload sorted App Sheet.\n";
                }
                break;
            }
            case 7: {
                string outApp;
                cout << "Enter export path for App Sheet: ";
                getline(cin, outApp);
                exportSheet(appSheet, outApp);
                break;
            }
            case 8:
                // Save all sheets to their original files
                writeCSV(stockPath, stockSheet);
                writeCSV(pricePath, priceSheet);
                writeCSV(appPath, appSheet);
                cout << "All sheets saved. Exiting.\n";
                return 0;
            case 9: {
                // Convert the latest App Sheet CSV (could be the synced one) to XLSX
                string appXlsxPath = appPath.substr(0, appPath.find_last_of('.')) + ".xlsx";
                vector<string> pythonCommands = {
                    "python html_to_csv_converter.py --csv-to-xlsx \"" + appPath + "\" \"" + appXlsxPath + "\"",
                    "python3 html_to_csv_converter.py --csv-to-xlsx \"" + appPath + "\" \"" + appXlsxPath + "\"",
                    "py html_to_csv_converter.py --csv-to-xlsx \"" + appPath + "\" \"" + appXlsxPath + "\""
                };
                bool converted = false;
                for (const string& cmd : pythonCommands) {
                    cout << "Trying command: " << cmd << endl;
                    int result = system(cmd.c_str());
                    if (result == 0) {
                        cout << "✓ App Sheet reconverted to: " << appXlsxPath << endl;
                        converted = true;
                        break;
                    }
                }
                if (!converted) {
                    cout << "✗ Failed to reconvert App Sheet to XLSX" << endl;
                }
                break;
            }
            case 10:
                deleteFirstNRowsFromPriceSheet(priceSheet, 10);
                break;
            case 11:
                unmergeIJColumnsInPriceSheet(priceSheet);
                break;
            case 12:
                movePColumnToHInPriceSheet(priceSheet);
                break;
            case 13:
                deleteRepeatedPriceColumn(priceSheet);
                break;
            case 14:
                deleteFirstSixColumns(priceSheet);
                break;
            case 15:
                setHomeNursingStockTo9000000(appSheet);
                break;
            case 16:
                convertNanToZero(appSheet);
                break;
            default:
                cout << "Invalid option. Try again.\n";
                break;
        }
    }
    return 0;
}
