#pragma once
#include <string>
#include <vector>

// Search for a product by code in a given sheet
typedef std::vector<std::vector<std::string>> SheetData;

void searchProduct(const std::string& code, const SheetData& stockSheet, const SheetData& priceSheet, const SheetData& appSheet);
void updateStock(SheetData& stockSheet, const std::string& code, int newStock);
void updatePrice(SheetData& priceSheet, const std::string& code, double newPrice);
void addProduct(SheetData& sheet, const std::vector<std::string>& productRow);
void syncAppSheet(SheetData& appSheet, const SheetData& stockSheet, const SheetData& priceSheet);
void sortAppSheet(SheetData& appSheet);
void removeRowsWithSkuContainingMultipleNumbers(SheetData& appSheet);
void exportSheet(const SheetData& sheet, const std::string& path);
void deleteFirstNRowsFromPriceSheet(SheetData& priceSheet, int n);
void unmergeIJColumnsInPriceSheet(SheetData& priceSheet);
void movePColumnToHInPriceSheet(SheetData& priceSheet);
void deleteRepeatedPriceColumn(SheetData& priceSheet);
void deleteFirstSixColumns(SheetData& priceSheet);
void setHomeNursingStockTo9000000(SheetData& appSheet);
void convertNanToZero(SheetData& appSheet);
void setMaxStockForSRProducts(SheetData& appSheet, const SheetData& stockSheet);
