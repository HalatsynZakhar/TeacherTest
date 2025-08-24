@echo off
:: 1. Create project directory in user folder (if not exists)
if not exist "%USERPROFILE%\ImageProcessor" mkdir "%USERPROFILE%\ImageProcessor"
:: 2. Navigate to project directory
cd /d "%USERPROFILE%\ImageProcessor"
:: 3. Check if repository folder already exists
if exist TeacherTest (
    echo TeacherTest folder already exists
    echo If you want a fresh installation, please delete or rename the existing folder
    pause
    exit
)
:: 4. Clone repository
git clone https://github.com/HalatsynZakhar/TeacherTest
:: 5. Navigate to project directory
cd TeacherTest
:: 6. Install dependencies
pip install -r requirements.txt
echo.
echo Project installation completed
pause