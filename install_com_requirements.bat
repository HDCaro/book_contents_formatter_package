@echo off
echo Installing Word COM Interface Requirements
echo ========================================
echo.

echo Installing pywin32 package...
pip install pywin32

echo.
echo Registering COM components...
python -c "import win32com.client; print('COM interface ready!')"

echo.
echo Setup complete! You can now run:
echo python extract_with_word_com.py
echo.
pause