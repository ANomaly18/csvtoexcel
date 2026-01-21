@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building Executable...
python -m PyInstaller --noconsole --onefile --name "CtE_Converter_v1_0_0_Final" --collect-all ttkbootstrap --collect-all tkinterdnd2 converter_app.py

echo.
echo Build Complete! Check the 'dist' folder.
pause
