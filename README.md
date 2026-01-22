# Excel → PDF Export Tool (Python, Tkinter)

This is a Windows desktop application that converts Excel files into a formatted PDF document using a custom Lao font (Saysettha).  
Supports 12×2 layout, auto column detection, and font embedding for PyInstaller builds.

## Features
- GUI built with Tkinter
- Auto-detects `Name` and `With` columns
- Uses embedded Saysettha fonts
- Exports PDF with proper 12×2 layout
- Compatible with PyInstaller EXE

## Run Locally

pipenv install
pipenv shell
python excel_to_pdf_app.py


## Build EXE

pyinstaller --noconsole --onefile ^
--add-data "Saysettha-Regular.ttf;." ^
--add-data "Saysettha-Bold.ttf;." ^
excel_to_pdf_app.py

Commit & push:

git add README.md
git commit -m "Add README"
git push
