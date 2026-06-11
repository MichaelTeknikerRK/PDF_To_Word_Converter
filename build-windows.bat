@echo off
REM Bygg "PDF til Word.exe" lokalt paa en Windows-maskin (alternativ til GitHub Actions).
REM Krever: Python 3.10+ og Tesseract installert (choco install tesseract / UB-Mannheim).
pip install -r requirements.txt pyinstaller
if not exist tesseract\tessdata mkdir tesseract\tessdata
copy "C:\Program Files\Tesseract-OCR\tesseract.exe" tesseract\ >nul
copy "C:\Program Files\Tesseract-OCR\*.dll" tesseract\ >nul
copy "C:\Program Files\Tesseract-OCR\tessdata\eng.traineddata" tesseract\tessdata\ >nul
if not exist tesseract\tessdata\nor.traineddata (
  curl -L -o tesseract\tessdata\nor.traineddata https://github.com/tesseract-ocr/tessdata_fast/raw/main/nor.traineddata
)
pyinstaller --onefile --windowed --name "PDF_To_Word_Converter" --add-data "tesseract;tesseract" --collect-data tkinterdnd2 app.py
echo.
echo Ferdig! Programmet ligger i dist\PDF_To_Word_Converter.exe
pause
