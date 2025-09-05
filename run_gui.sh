#!/bin/bash
# Launcher script for OFT to EML Converter GUI

# Get the directory of the script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# Activate virtual environment and run the GUI
cd "$DIR"

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Setting up virtual environment..."
    python -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

echo "Starting OFT to EML Converter GUI..."
python oft_to_eml_gui.py