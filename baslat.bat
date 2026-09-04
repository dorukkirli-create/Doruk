@echo off
REM ============================================================
REM  Masraf Merkezi Otomasyonu - Windows baslatici
REM  Cift tiklayarak calistirin. Ilk calistirmada sanal ortam
REM  kurulur ve paketler indirilir; sonraki acilislar hizlidir.
REM ============================================================
setlocal
cd /d "%~dp0"

echo ============================================
echo   Masraf Merkezi Otomasyonu baslatiliyor...
echo ============================================
echo.

REM --- Python yorumlayicisini bul -----------------------------
REM Once 'py' baslaticisi denenir: Windows kurulumlarinda standarttir
REM ve Microsoft Store kisayolu sorununu yasamaz. Sonra 'python'.
REM DIKKAT: Windows'ta yonlendirme 'nul' olur, '/dev/null' DEGIL.
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python --version >nul 2>&1 && set "PY=python"
)

if not defined PY (
    echo HATA: Python bulunamadi.
    echo.
    echo Cozum: python.org/downloads adresinden Python 3.11 veya
    echo uzerini kurun. Kurulum ekranindaki
    echo    [x] Add Python to PATH
    echo kutusunu MUTLAKA isaretleyin.
    echo.
    pause
    exit /b 1
)

echo Python bulundu: %PY%
%PY% --version
echo.

REM --- Sanal ortam --------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Sanal ortam olusturuluyor, bu islem yalnizca bir kez yapilir...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo.
        echo HATA: Sanal ortam olusturulamadi.
        echo Klasorde yazma izniniz olmayabilir. Klasoru Belgeler
        echo altina tasiyip tekrar deneyin.
        pause
        exit /b 1
    )
)

set "VENV_PY=.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo HATA: Sanal ortam bozuk gorunuyor.
    echo .venv klasorunu silip bu dosyayi tekrar calistirin.
    pause
    exit /b 1
)

REM --- Paketler -----------------------------------------------
REM Streamlit zaten kuruluysa internete cikmayalim; her acilista
REM kullaniciyi bekletmenin anlami yok.
"%VENV_PY%" -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Gerekli paketler kuruluyor, birkac dakika surebilir...
    "%VENV_PY%" -m pip install --quiet --upgrade pip
    "%VENV_PY%" -m pip install --quiet -r requirements.txt
    if errorlevel 1 (
        echo.
        echo HATA: Paketler kurulamadi.
        echo Internet baglantisini veya sirket proxy ayarlarini kontrol edin.
        pause
        exit /b 1
    )
    echo Paketler kuruldu.
) else (
    echo Paketler hazir.
)

echo.
echo ============================================
echo   Arayuz aciliyor: http://localhost:8501
echo   Kapatmak icin bu pencerede Ctrl+C
echo ============================================
echo.

"%VENV_PY%" -m streamlit run app.py

echo.
echo Uygulama kapandi.
pause
