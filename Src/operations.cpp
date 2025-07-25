#include "operations.h"
#include "file_handler.h"
#include "logger.h"
#include <iostream>
#include <algorithm>
#include <cctype>
#include <regex>
#include <limits>
using namespace std;

#pragma once
#include <iostream>
#include <vector>
#include <string>
#include <algorithm>

using namespace std;
using SheetData = vector<vector<string>>;

// Utility to trim whitespace and hidden characters from strings
string trim(const string& str) {
    const char* ws = " \t\n\r\f\v";
    size_t start = str.find_first_not_of(ws);
    size_t end = str.find_last_not_of(ws);
    return (start == string::npos) ? "" : str.substr(start, end - start + 1);
}

// Find a column by header name in first row
int findColumnByName(const SheetData& sheet, const string& headerName) {
    if (sheet.empty()) return -1;
    for (size_t i = 0; i < sheet[0].size(); ++i) {
        if (trim(sheet[0][i]) == trim(headerName)) return static_cast<int>(i);
    }
    return -1;
}

// Find a row by exact code match in a given column
int findRowByCode(const SheetData& sheet, const string& code, int codeCol = 0) {
    for (size_t i = 0; i < sheet.size(); ++i) {
        if (sheet[i].size() > codeCol && trim(sheet[i][codeCol]) == trim(code))
            return static_cast<int>(i);
    }
    return -1;
}

// Helper: Normalize all rows to the header length
void normalizeAppSheetRowsToHeader(SheetData& appSheet) {
    if (appSheet.empty()) return;
    size_t headerLen = appSheet[0].size();
    for (size_t i = 1; i < appSheet.size(); ++i) {
        if (appSheet[i].size() < headerLen) {
            appSheet[i].resize(headerLen, ""); // pad with blanks
        } else if (appSheet[i].size() > headerLen) {
            appSheet[i].resize(headerLen); // truncate extra columns
        }
    }
}

// Search for a product in all sheets using integer indices
void searchProduct(const string& code, const SheetData& stockSheet, const SheetData& priceSheet, const SheetData& appSheet) {
    cout << "\n--- Search Results for code: " << code << " ---\n";

    if (!stockSheet.empty()) {
        int codeCol = 0; // code column in stockSheet
        int totalCol = 21; // total column in stockSheet
        int rowIdx = findRowByCode(stockSheet, code, codeCol);

        if (rowIdx != -1 && totalCol < (int)stockSheet[rowIdx].size()) {
            cout << "Stock Sheet: Row " << rowIdx << "\n";
            for (const auto& cell : stockSheet[rowIdx]) cout << cell << ", ";
            cout << "\nTotal = " << stockSheet[rowIdx][totalCol] << "\n";
        } else {
            cout << "Stock Sheet: Not found\n";
        }
    }

    if (!priceSheet.empty()) {
        int codeCol = 15; // الكود column in priceSheet
        int priceCol = 8; // السعر column in priceSheet
        int rowIdx = findRowByCode(priceSheet, code, codeCol);

        if (rowIdx != -1 && priceCol < (int)priceSheet[rowIdx].size()) {
            cout << "Price Sheet: Row " << rowIdx << "\n";
            for (const auto& cell : priceSheet[rowIdx]) cout << cell << ", ";
            cout << "\nPrice = " << priceSheet[rowIdx][priceCol] << "\n";
        } else {
            cout << "Price Sheet: Not found\n";
        }
    }

    if (!appSheet.empty()) {
        int skuCol = 8; // sku column in appSheet
        int priceCol = 9; // price column in appSheet
        int stockCol = 27; // stock column in appSheet
        int rowIdx = findRowByCode(appSheet, code, skuCol);

        if (rowIdx != -1 && priceCol < (int)appSheet[rowIdx].size() && stockCol < (int)appSheet[rowIdx].size()) {
            cout << "App Sheet: Row " << rowIdx << "\n";
            for (const auto& cell : appSheet[rowIdx]) cout << cell << ", ";
            cout << "\nPrice = " << appSheet[rowIdx][priceCol] << ", Stock = " << appSheet[rowIdx][stockCol] << "\n";
        } else {
            cout << "App Sheet: Not found\n";
        }
    }
}

// Update stock in stock sheet using your specified index
void updateStock(SheetData& stockSheet, const string& code, int newStock) {
    int codeCol = 0; // code column
    int totalCol = 21; // total column
    int idx = findRowByCode(stockSheet, code, codeCol);
    if (idx != -1 && totalCol < (int)stockSheet[idx].size()) {
        stockSheet[idx][totalCol] = to_string(newStock);
        logChange("Stock updated for code " + code + ": new stock = " + to_string(newStock));
        cout << "Stock updated.\n";
    } else {
        cout << "Product not found in stock sheet.\n";
    }
}

// Update price in price sheet using integer indices
void updatePrice(SheetData& priceSheet, const string& code, double newPrice) {
    int codeCol = 14; // الكود column
    int priceCol = 8; // السعر column
    int idx = findRowByCode(priceSheet, code, codeCol);
    if (idx != -1 && priceCol < (int)priceSheet[idx].size()) {
        priceSheet[idx][priceCol] = to_string(newPrice);
        logChange("Price updated for الكود " + code + ": new السعر = " + to_string(newPrice));
        cout << "Price updated.\n";
    } else {
        cout << "Product not found in price sheet.\n";
    }
}

// Add a new product to a sheet (no change needed for indices)
void addProduct(SheetData& sheet, const vector<string>& productRow) {
    sheet.push_back(productRow);
    logChange("Product added: code = " + productRow[0]);
    cout << "Product added.\n";
}

// Synchronize App Sheet with Stock and Price Sheets using VLOOKUP-like logic
void syncAppSheet(SheetData& appSheet, const SheetData& stockSheet, const SheetData& priceSheet) {
    if (appSheet.empty() || stockSheet.empty() || priceSheet.empty()) {
        cout << "Error: One or more sheets are empty. Cannot synchronize.\n";
        return;
    }

    // Find column indices dynamically
    int priceCodeCol = 14; // الكود column in pricesheet
    int priceValueCol = 8; // السعر column in pricesheet
    int stockCodeCol = 0; // code column in stockSheet
    int stockTotalCol = 21; // total column in stockSheet
    int appSkuCol = 8; // sku column in appSheet
    int appPriceCol = 9; // price column in appSheet
    int appStockCol = 27; // stock column in appSheet

    if (priceCodeCol == -1 || priceValueCol == -1) {
        cout << "Error: Could not find الكود or السعر columns in Price Sheet.\n";
        return;
    }

    int updatedCount = 0;
    for (size_t i = 1; i < appSheet.size(); ++i) { // skip header
        if (appSheet[i].size() <= max(appSkuCol, max(appPriceCol, appStockCol))) continue;
        string sku = trim(appSheet[i][appSkuCol]);
        if (sku.empty()) continue;

        // VLOOKUP in Price Sheet
        int priceRow = findRowByCode(priceSheet, sku, priceCodeCol);
        if (priceRow != -1 && priceValueCol < (int)priceSheet[priceRow].size()) {
            string newPrice = trim(priceSheet[priceRow][priceValueCol]);
            if (!newPrice.empty() && newPrice != appSheet[i][appPriceCol]) {
                appSheet[i][appPriceCol] = newPrice;
                updatedCount++;
            }
        }

        // VLOOKUP in Stock Sheet
        int stockRow = findRowByCode(stockSheet, sku, stockCodeCol);
        if (stockRow != -1 && stockTotalCol < (int)stockSheet[stockRow].size()) {
            string newStock = trim(stockSheet[stockRow][stockTotalCol]);
            if (!newStock.empty() && newStock != appSheet[i][appStockCol]) {
                appSheet[i][appStockCol] = newStock;
                updatedCount++;
            }
        }
    }
    logChange("App Sheet synchronized using VLOOKUP from Stock and Price Sheets - " + to_string(updatedCount) + " cells updated.");
    cout << "App Sheet synchronized. Updated " << updatedCount << " cells.\n";
    // Call the new function for #S#R products
    setMaxStockForSRProducts(appSheet, stockSheet);
}

/**
void sortAppSheet(SheetData& appSheet) {
    if (appSheet.size() <= 1) {
        cout << "App Sheet is empty or has only header row. Nothing to sort.\n";
        return;
    }
    int skuCol = 8; // or use dynamic detection if needed

    // Clean all SKUs: remove all non-digit characters
    for (size_t i = 1; i < appSheet.size(); ++i) {
        if (appSheet[i].size() > skuCol) {
            string& s = appSheet[i][skuCol];
            string digits;
            for (char c : s) if (isdigit(c)) digits += c;
            s = digits;
        }
    }

    // Now sort numerically, treating blanks as very large numbers
    sort(appSheet.begin() + 1, appSheet.end(), [skuCol](const vector<string>& a, const vector<string>& b) {
        double skuA = (a.size() > skuCol && !a[skuCol].empty()) ? stod(a[skuCol]) : std::numeric_limits<double>::max();
        double skuB = (b.size() > skuCol && !b[skuCol].empty()) ? stod(b[skuCol]) : std::numeric_limits<double>::max();
        return skuA < skuB;
    });

    logChange("App Sheet sorted by sku (final robust numeric, header preserved).");
    cout << "App Sheet sorted in ascending order by sku (final robust numeric, header row preserved).\n";
}

// Remove any row where SKU column contains two or more separate numbers (sequences of digits)
void removeRowsWithSkuContainingMultipleNumbers(SheetData& appSheet) {
    int skuCol = 8; // sku column in appSheet
    // Normalize all rows to header length to prevent shifting/corruption
    normalizeAppSheetRowsToHeader(appSheet);
    std::regex numberPattern(R"(\d+)");
    auto hasMultipleNumbers = [skuCol, &numberPattern](const vector<string>& row) {
        if (row.size() <= skuCol) return false;
        auto sku = row[skuCol];
        auto numbers_begin = std::sregex_iterator(sku.begin(), sku.end(), numberPattern);
        auto numbers_end = std::sregex_iterator();
        int count = std::distance(numbers_begin, numbers_end);
        return count >= 2;
    };
    size_t before = appSheet.size();
    // Only erase from data rows, not header
    appSheet.erase(remove_if(appSheet.begin() + 1, appSheet.end(), hasMultipleNumbers), appSheet.end());
    size_t after = appSheet.size();
    if (before != after) {
        logChange("Rows with multiple numbers in SKU removed from App Sheet.");
        cout << (before - after) << " row(s) with multiple numbers in SKU removed from App Sheet.\n";
    } else {
        cout << "No rows with multiple numbers in SKU removed from App Sheet.\n";
    }
}
**/

// Export a sheet to a new file
void exportSheet(const SheetData& sheet, const string& path) {
    if (writeCSV(path, sheet)) {
        cout << "Sheet exported to " << path << endl;
    } else {
        cout << "Failed to export sheet to " << path << endl;
    }
}

// Deletes the first N rows from the price sheet (excluding header)
void deleteFirstNRowsFromPriceSheet(SheetData& priceSheet, int n) {
    if (priceSheet.size() == 0) {
        cout << "Price sheet is empty.\n";
        return;
    }
    int toDelete = min(n, (int)priceSheet.size());
    priceSheet.erase(priceSheet.begin(), priceSheet.begin() + toDelete);
    logChange("Deleted first " + to_string(toDelete) + " rows from price sheet (including header).");
    cout << "Deleted first " << toDelete << " rows from price sheet.\n";
}

// Unmerges columns I & J in the price sheet
void unmergeIJColumnsInPriceSheet(SheetData& priceSheet) {
    // Columns I (8) and J (9) (0-based)
    for (size_t i = 1; i < priceSheet.size(); ++i) {
        if (priceSheet[i].size() > 9) {
            // If I and J are merged (e.g., "val1|val2"), split by '|' or other delimiter
            string& colI = priceSheet[i][8];
            string& colJ = priceSheet[i][9];
            size_t delim = colI.find('|');
            if (delim != string::npos) {
                colJ = colI.substr(delim + 1);
                colI = colI.substr(0, delim);
            }
        }
    }
    logChange("Unmerged columns I & J in price sheet.");
    cout << "Unmerged columns I & J in price sheet.\n";
}

// Moves the P column (referring to الكود) to the H column in the price sheet
void movePColumnToHInPriceSheet(SheetData& priceSheet) {
    // Column P (15), H (7) (0-based)
    for (size_t i = 1; i < priceSheet.size(); ++i) {
        if (priceSheet[i].size() > 15 && priceSheet[i].size() > 7) {
            priceSheet[i][7] = priceSheet[i][15];
        }
    }
    logChange("Moved P column (الكود) to H column in price sheet.");
    cout << "Moved P column (الكود) to H column in price sheet.\n";
}

// Delete the repeated السعر column (index 9) from the price sheet
void deleteRepeatedPriceColumn(SheetData& priceSheet) {
    int priceColToDelete = 9; // السعر column to delete

    for (size_t i = 0; i < priceSheet.size(); ++i) {
        if (priceSheet[i].size() > priceColToDelete) {
            priceSheet[i].erase(priceSheet[i].begin() + priceColToDelete);
        }
    }
    logChange("Deleted repeated السعر column (index 9) from price sheet.");
    cout << "Deleted repeated السعر column (index 9) from price sheet.\n";
}

// Delete the first 6 columns from the price sheet
void deleteFirstSixColumns(SheetData& priceSheet) {
    int columnsToDelete = 6; // Number of columns to delete from the beginning

    for (size_t i = 0; i < priceSheet.size(); ++i) {
        if (priceSheet[i].size() > columnsToDelete) {
            priceSheet[i].erase(priceSheet[i].begin(), priceSheet[i].begin() + columnsToDelete);
        } else if (priceSheet[i].size() > 0) {
            // If row has fewer columns than 6, delete all columns
            priceSheet[i].clear();
        }
    }
    logChange("Deleted first 6 columns from price sheet.");
    cout << "Deleted first 6 columns from price sheet.\n";
}

// Set all stock values for "home nursing services" to 9000000 in App Sheet
void setHomeNursingStockTo9000000(SheetData& appSheet) {
    int stockCol = 27; // stock column in appSheet
    int updatedCount = 0;

    for (size_t i = 1; i < appSheet.size(); ++i) { // skip header, start from row 1
        if (appSheet[i].size() > stockCol) {
            // Check if this row contains "home nursing services" (case insensitive)
            bool isHomeNursing = false;
            for (const string& cell : appSheet[i]) {
                string lowerCell = cell;
                transform(lowerCell.begin(), lowerCell.end(), lowerCell.begin(), ::tolower);
                if (lowerCell.find("home nursing services") != string::npos) {
                    isHomeNursing = true;
                    break;
                }
            }

            if (isHomeNursing) {
                appSheet[i][stockCol] = "9000000";
                updatedCount++;
            }
        }
    }

    logChange("Set stock to 9000000 for " + to_string(updatedCount) + " home nursing services rows in App Sheet.");
    cout << "Updated " << updatedCount << " home nursing services rows with stock = 9000000.\n";
}

// Convert all "nan" values to "0" in App Sheet
void convertNanToZero(SheetData& appSheet) {
    int convertedCount = 0;

    for (size_t i = 0; i < appSheet.size(); ++i) {
        for (size_t j = 0; j < appSheet[i].size(); ++j) {
            string& cell = appSheet[i][j];
            string lowerCell = cell;
            transform(lowerCell.begin(), lowerCell.end(), lowerCell.begin(), ::tolower);

            // Check for various forms of "nan"
            if (lowerCell == "nan" || lowerCell == "n/a" || lowerCell == "na" ||
                lowerCell == "null" || lowerCell == "undefined" || lowerCell == "") {
                cell = "0";
                convertedCount++;
            }
        }
    }

    logChange("Converted " + to_string(convertedCount) + " nan/empty values to 0 in App Sheet.");
    cout << "Converted " << convertedCount << " nan/empty values to 0 in App Sheet.\n";
}

// Set max stock quantity in app sheet (column 25) for products with #S#R in their code
void setMaxStockForSRProducts(SheetData& appSheet, const SheetData& stockSheet) {
    int stockCodeCol = 0; // code column in stockSheet
    int appSkuCol = 8;    // sku column in appSheet
    int maxStockCol = 25; // max stock column in appSheet
    int updatedCount = 0;
    // Collect all codes in stockSheet with #S#R
    std::vector<std::string> srCodes;
    for (size_t i = 1; i < stockSheet.size(); ++i) { // skip header
        if (stockSheet[i].size() > stockCodeCol) {
            const std::string& code = stockSheet[i][stockCodeCol];
            if (code.find("#S#R") != std::string::npos) {
                srCodes.push_back(trim(code));
            }
        }
    }
    // For each code, find the matching row in appSheet and set max stock
    for (const auto& code : srCodes) {
        for (size_t i = 1; i < appSheet.size(); ++i) {
            if (appSheet[i].size() > maxStockCol && appSheet[i].size() > appSkuCol) {
                if (trim(appSheet[i][appSkuCol]) == code) {
                    appSheet[i][maxStockCol] = "1";
                    updatedCount++;
                }
            }
        }
    }
    logChange("Set max stock for #S#R products in App Sheet for " + std::to_string(updatedCount) + " rows.");
    std::cout << "Updated max stock for #S#R products in App Sheet for " << updatedCount << " rows.\n";
}
