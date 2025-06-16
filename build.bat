@echo off
echo === Building Work Schedule Generator Executable ===

:: Kiểm tra PyInstaller
echo Checking for PyInstaller...
where pyinstaller >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

:: Kiểm tra và cài đặt openpyxl
echo Checking dependencies...
python -c "import openpyxl" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo openpyxl not found. Installing...
    pip install openpyxl
)

:: Tạo thư mục dist nếu chưa có
if not exist dist mkdir dist

echo Building for Windows...

:: Build executable với PyInstaller
pyinstaller --onefile ^
    --name="WorkScheduleGenerator" ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=csv ^
    --hidden-import=glob ^
    --hidden-import=datetime ^
    --hidden-import=collections ^
    --clean ^
    scd3.py

echo Build completed!

:: Kiểm tra kết quả build
if exist "dist\WorkScheduleGenerator.exe" (
    echo ✅ Build successful!
    echo 📁 Executable file: dist\WorkScheduleGenerator.exe
    dir "dist\WorkScheduleGenerator.exe"
) else (
    echo ❌ Build failed!
    exit /b 1
)
