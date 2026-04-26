#!/bin/bash
# Extract Reports — Bear Engineering
# Double-click this file to run the report extractor.

# Change to the folder where this script lives
cd "$(dirname "$0")"

echo "======================================"
echo "  Bear Engineering — Report Extractor"
echo "======================================"
echo ""

python3 extract_reports.py

echo ""
echo "Press any key to close this window..."
read -n 1 -s
