#!/bin/bash
# Mac launcher — double-click to run

# Change to the folder where this script lives
cd "$(dirname "$0")"

echo "Installing dependencies..."
pip3 install -r requirements.txt --quiet

echo ""
echo "Starting app — your browser will open automatically."
echo "Press Ctrl+C in this window to stop the app."
echo ""
streamlit run app.py
