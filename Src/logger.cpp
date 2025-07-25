#include "logger.h"
#include <fstream>
#include <ctime>
#include <iostream>
using namespace std;

// Logs a change message to a log file with a timestamp
void logChange(const string& message) {
    ofstream logFile("log.txt", ios::app);
    if (!logFile.is_open()) {
        cerr << "Error: Could not open log file." << endl;
        return;
    }
    // Get current time
    time_t now = time(0);
    char* dt = ctime(&now);
    // Remove newline from ctime
    string timeStr(dt);
    if (!timeStr.empty() && timeStr[timeStr.length()-1] == '\n')
        timeStr.erase(timeStr.length()-1);
    logFile << "[" << timeStr << "] " << message << endl;
    logFile.close();
}
