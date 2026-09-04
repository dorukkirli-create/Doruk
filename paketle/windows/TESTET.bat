@echo off
REM ============================================================
REM  KURULUM TESTI
REM  Bir sey calismadiginda bu dosyaya cift tiklayin ve ciktinin
REM  tamamini kopyalayip iletin. Hicbir sey degistirmez, sadece bakar.
REM ============================================================
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"

set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONSTARTUP="
set "PYTHONEXECUTABLE="
set "PYTHONUSERBASE="
set "PYTHONNOUSERSITE=1"

echo ============================================================
echo   KURULUM TESTI
echo ============================================================
echo.
echo Klasor: %~dp0
echo.

echo --- 1. Klasor yapisi ---------------------------------------
for %%D in (1_FATURALAR 2_EXCEL_CIKTI PERSONEL program veri) do (
    if exist "%%D\" (echo   VAR    %%D) else (echo   EKSIK  %%D)
)
for %%F in (program\python.exe program\python311.dll program\Lib\os.py program\kod\calistir.py) do (
    if exist "%%F" (echo   VAR    %%F) else (echo   EKSIK  %%F)
)
echo.

echo --- 2. Sizdiran Python ortam degiskenleri ------------------
echo   PYTHONHOME = [%PYTHONHOME%]
echo   PYTHONPATH = [%PYTHONPATH%]
echo   (Kose parantezler bos olmali)
echo.

echo --- 3. Eski kalintilar -------------------------------------
if exist ".venv\" (
    echo   DIKKAT: .venv klasoru var. Eski kurulumdan kalmis olabilir,
    echo           silmek guvenlidir.
) else (
    echo   temiz  .venv yok
)
if exist "baslat.bat" (
    echo   DIKKAT: baslat.bat var. Bu ESKI baslaticidir ve Python kurulumu
    echo           ister. Bu paketle kullanilmaz, silebilirsiniz.
) else (
    echo   temiz  baslat.bat yok
)
echo.

echo --- 4. Gomulu Python calisiyor mu --------------------------
if not exist "program\python.exe" (
    echo   BASARISIZ: program\python.exe yok.
    goto :son
)
"program\python.exe" -c "import sys; print('   surum   :', sys.version.split()[0]); print('   konum   :', sys.executable); print('   kutuphane:', sys.prefix)"
if errorlevel 1 (
    echo   BASARISIZ: Python baslatilamadi. Yukaridaki hata mesaji onemli.
    goto :son
)
echo.

echo --- 5. Gerekli kutuphaneler --------------------------------
"program\python.exe" -c "import importlib,sys; [print(('   VAR   ' if importlib.util.find_spec(m) else '   EKSIK ') + m) for m in ('pandas','numpy','openpyxl','xlrd','rapidfuzz','xlsxwriter','extract_msg','olefile','tzlocal','tzdata','red_black_dict_mod')]"
echo.

echo --- 6. Saat dilimi verisi (tzdata) -------------------------
"program\python.exe" -c "from zoneinfo import ZoneInfo; ZoneInfo('Europe/Moscow'); print('   TAMAM: saat dilimi verisi okunuyor')" 2>&1
echo.

echo --- 7. Uygulama modulu -------------------------------------
"program\python.exe" -c "import sys; sys.path.insert(0,'program/kod'); import masraf.boru; print('   TAMAM: masraf.boru ice aktarildi')" 2>&1
echo.

echo --- 8. Personel dosyalari ----------------------------------
if exist "PERSONEL\*.xlsx" (
    for %%F in (PERSONEL\*.xlsx) do echo   VAR    %%~nxF  (%%~zF bayt)
) else (
    echo   EKSIK: PERSONEL klasorunde .xlsx dosyasi yok
)
echo.

echo --- 9. Fatura dosyalari ------------------------------------
set SAYI=0
for %%F in (1_FATURALAR\*.msg 1_FATURALAR\*.xlsx 1_FATURALAR\*.xls) do set /a SAYI+=1
echo   1_FATURALAR icinde islenebilir dosya sayisi: %SAYI%

:son
echo.
echo ============================================================
echo   TEST BITTI. Bu ciktinin TAMAMINI kopyalayip iletin.
echo ============================================================
pause
