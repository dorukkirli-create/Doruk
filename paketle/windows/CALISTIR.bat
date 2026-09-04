@echo off
REM ============================================================
REM  MASRAF MERKEZI OTOMASYONU
REM
REM  Kullanim:
REM    1) Faturalari 1_FATURALAR klasorune atin
REM    2) Bu dosyaya cift tiklayin
REM  ya da dosyalari dogrudan bu dosyanin uzerine surukleyip birakin.
REM
REM  Python KURMANIZA GEREK YOKTUR. Gereken her sey program klasorunde.
REM ============================================================
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"

REM --- Miras alinan Python degiskenlerini TEMIZLE ---------------
REM Makinede baska bir Python varsa PYTHONHOME/PYTHONPATH tanimli olabilir.
REM Bunlar gomulu yorumlayiciyi kendi kutuphanesi yerine baska bir yere
REM yonlendirir ve 'No pyvenv.cfg file' gibi hatalara yol acar.
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONSTARTUP="
set "PYTHONEXECUTABLE="
set "PYTHONUSERBASE="
set "PYTHONNOUSERSITE=1"

set "PYEXE=%~dp0program\python.exe"

if not exist "%PYEXE%" (
    echo.
    echo HATA: program\python.exe bulunamadi.
    echo.
    echo Klasor eksik ya da ZIP tam acilmamis olabilir.
    echo ZIP dosyasini tekrar acin ve TUM Otomasyon klasorunu kopyalayin.
    echo.
    pause
    exit /b 1
)

REM --- Yorumlayici gercekten calisiyor mu? ----------------------
REM Dosyanin var olmasi yetmez; kutuphanesi eksikse baslamaz.
"%PYEXE%" -c "import sys, os, zipfile" >nul 2>&1
if errorlevel 1 (
    echo.
    echo HATA: Pakete gomulu Python baslatilamadi.
    echo.
    echo Ayrinti icin asagidaki mesaji okuyun:
    echo ------------------------------------------------------------
    "%PYEXE%" -c "import sys; print(sys.version)"
    echo ------------------------------------------------------------
    echo.
    echo En sik sebep: ZIP eksik acilmis. Cozum:
    echo   1. Otomasyon klasorunu tamamen silin
    echo   2. ZIP dosyasina sag tiklayip "Tumunu ayikla" deyin
    echo   3. Cikan klasoru masaustune kopyalayin
    echo.
    echo Sorun surerse TESTET.bat dosyasini calistirip ciktiyi iletin.
    echo.
    pause
    exit /b 1
)

"%PYEXE%" "%~dp0program\kod\calistir.py" %*
set "SONUC=%ERRORLEVEL%"

echo.
if not "%SONUC%"=="0" (
    echo Islem hatayla bitti. Yukaridaki mesaji okuyun.
)
echo Kapatmak icin bir tusa basin.
pause >nul
exit /b %SONUC%
