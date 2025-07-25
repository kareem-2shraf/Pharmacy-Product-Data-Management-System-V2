#include "file_handler.h"
#include <fstream>
#include <sstream>
#include <iostream>
using namespace std;

// Reads a CSV file into a 2D vector. Returns true if successful.
bool readCSV(const string& path, vector<vector<string>>& data) {
    ifstream file(path.c_str());
    if (!file.is_open()) {
        cerr << "Error: Could not open file " << path << endl;
        return false;
    }
    data.clear();
    string line;
    while (getline(file, line)) {
        data.push_back(splitCSVLine(line));
    }
    file.close();
    return true;
}

// Writes a 2D vector to a CSV file. Returns true if successful.
bool writeCSV(const string& path, const vector<vector<string>>& data) {
    ofstream file(path.c_str());
    if (!file.is_open()) {
        cerr << "Error: Could not open file for writing: " << path << endl;
        return false;
    }
    for (size_t i = 0; i < data.size(); ++i) {
        file << joinCSVLine(data[i]);
        if (i != data.size() - 1) file << '\n';
    }
    file.close();
    return true;
}

// Splits a CSV line into fields.
vector<string> splitCSVLine(const string& line) {
    vector<string> result;
    stringstream ss(line);
    string item;
    while (getline(ss, item, ',')) {
        result.push_back(item);
    }
    return result;
}

// Joins fields into a CSV line.
string joinCSVLine(const vector<string>& fields) {
    string line;
    for (size_t i = 0; i < fields.size(); ++i) {
        line += fields[i];
        if (i != fields.size() - 1) line += ",";
    }
    return line;
}
