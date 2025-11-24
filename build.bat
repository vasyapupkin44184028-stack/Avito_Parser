@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building EXE file...
pyinstaller --onefile --windowed --name "AvitoParserPro" --icon=icon.ico --add-data "chromedriver.exe;." avito_parser.py

echo Build complete!
pause