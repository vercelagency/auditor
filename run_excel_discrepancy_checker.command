#!/bin/bash
cd "$(dirname "$0")"

# Check for required packages
python3 -c "import streamlit, pandas, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
  echo "Installing required Python packages..."
  pip3 install -r requirements.txt
fi

nohup python3 -m streamlit run app.py &> streamlit.log & 