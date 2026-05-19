@echo off
setlocal

python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --clean --noconfirm SendEmailsTool.spec

echo.
if exist dist\SendEmailsTool\app.ico (
    echo app.ico included: dist\SendEmailsTool\app.ico
) else (
    echo WARNING: app.ico was not found in dist\SendEmailsTool
)
echo.
echo Build finished. Run dist\SendEmailsTool\SendEmailsTool.exe
