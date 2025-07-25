#pragma once
#include <vector>
#include <string>

// Reads a CSV file into a 2D vector. Returns true if successful.
bool readCSV(const std::string& path, std::vector<std::vector<std::string>>& data);

// Writes a 2D vector to a CSV file. Returns true if successful.
bool writeCSV(const std::string& path, const std::vector<std::vector<std::string>>& data);

// Splits a CSV line into fields.
std::vector<std::string> splitCSVLine(const std::string& line);

// Joins fields into a CSV line.
std::string joinCSVLine(const std::vector<std::string>& fields);
