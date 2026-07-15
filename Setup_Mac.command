#!/bin/bash
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
echo "Setup complete. Double-click Start_Maverick_Roof_Tool.command."
read -r -p "Press Return to close."
