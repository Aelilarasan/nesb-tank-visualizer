@echo off
cd /d C:\TankDashboard
:loop
git status --porcelain | findstr /R "." >nul
if %errorlevel% equ 0 (
    echo Changes detected in Excel! Syncing to GitHub...
    git add tank_data.xlsx
    git commit -m "Auto-update tank data"
    git push origin main
    echo Sync complete!
)
timeout /t 10 /nobreak >nul
goto loop