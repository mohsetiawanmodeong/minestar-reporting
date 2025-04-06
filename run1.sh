#!/bin/bash

# Check if Python is installed
if ! command -v python &> /dev/null && ! command -v python3 &> /dev/null; then
    echo "Error: Python is not installed. Please install Python first."
    exit 1
fi

# Use python3 if available, otherwise use python
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
else
    PYTHON_CMD="python"
fi

# Check if required packages are installed
if ! $PYTHON_CMD -c "import flask, pandas, openpyxl" &> /dev/null; then
    echo "Installing required packages..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

# Run the application
echo "Starting Excel Data Cleaner application..."
echo "Once started, open a web browser and go to http://127.0.0.1:5051"
$PYTHON_CMD app1.py 