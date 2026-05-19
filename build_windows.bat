@echo off
setlocal

python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --clean --noconfirm SendEmailsTool.spec

echo.
echo Build finished. Run dist\SendEmailsTool\SendEmailsTool.exe
