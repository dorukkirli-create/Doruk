@echo off
REM Masraf Merkezi Otomasyonu - Windows baslatici
REM Cift tiklayarak calistirin. Ilk calistirmada sanal ortam kurulur.
setlocal
cd /d "%~dp0"

echo ============================================
echo   Masraf Merkezi Otomasyonu baslatiliyor...
echo ============================================
echo.

where python >/dev/null 2>&1
if errorlevel 1 (
    echo HATA: Python bulunamadi.
    echo Lutfen python.org adresinden Python 3.11 kurun ve
    echo kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
    pause
    exit /b 1
)

if not exist .venv (
    echo Sanal ortam olusturuluyor, bu islem bir kez yapilir...
    python -m venv .venv
    if errorlevel 1 (
        echo HATA: Sanal ortam olusturulamadi.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

echo Gerekli paketler kontrol ediliyor...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt
if errorlevel 1 (
    echo HATA: Paketler kurulamadi. Internet baglantisini kontrol edin.
    echo Paketler daha once kurulduysa internet gerekmez, devam edebilirsiniz.
)

echo.
echo Arayuz aciliyor. Tarayici otomatik acilmazsa: http://localhost:8501
echo Kapatmak icin bu pencerede Ctrl+C tuslarina basin.
echo.
python -m streamlit run app.py

pause
