#!/bin/bash
# One-click launcher for RFID Attendance Dashboard & Excel Sync
cd "$(dirname "$0")"

echo "=================================================="
echo " Starting RFID Attendance Dashboard & Excel Sync"
echo "=================================================="

# Open the dashboard in default browser after 1 second
(sleep 1 && open "http://localhost:5050") &

# Run Python dashboard using PlatformIO virtualenv
~/.platformio/penv/bin/python attendance_dashboard.py

