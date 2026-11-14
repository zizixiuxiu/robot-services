@echo off
rem iShare sync trigger on port 8016 (triggered mode, no polling)
cd /d "D:\Services\robot-services\workshop-order"
"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe" "D:\Services\robot-services\workshop-order\scripts\ishare_sync_api.py"
