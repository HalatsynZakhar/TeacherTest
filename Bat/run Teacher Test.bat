@echo off

:: 1. Check if project directory exists
if not exist "%USERPROFILE%\ImageProcessor\TeacherTest" (
    echo Project directory not found
    echo Please run setup.bat first
    pause
    exit
)

:: 2. Navigate to project directory
cd /d "%USERPROFILE%\ImageProcessor\TeacherTest"

:: 3. Ensure we are on the correct branch
git checkout main

:: 4. Fetch latest changes
git fetch origin

:: 5. Remove untracked files and directories
git clean -fd

:: 6. Hard reset to remote branch
git reset --hard origin/main

:: 7. Launch project
python start.py

pause
