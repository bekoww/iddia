@echo off
chcp 65001 >nul
title Iddaa Oran Eslestirici
cd /d "%~dp0"
color 0A
mode con: cols=65 lines=22 >nul 2>nul
echo.
echo  =============================================
echo      Iddaa Oran Eslestirici
echo  =============================================
echo.
echo   [1] Tahmin Uygulamasini Baslat
echo   [2] Bugunun Oynanmamis Maclarini Excel'e Aktar
echo   [3] Gecmis Mac Verisi Topla
echo   [4] 2026 Verisini Kolay Topla
echo   [5] Excel Verilerini Ice Aktar
echo   [6] Projeyi Kontrol Et
echo   [7] Guncel Veriyi Excel'e Aktar
echo   [8] Cikis
echo.
set /p secim=  Seciminiz (1/2/3/4/5/6/7/8): 
if "%secim%"=="1" goto ui
if "%secim%"=="2" goto today_scraper
if "%secim%"=="3" goto scraper
if "%secim%"=="4" goto yearly_scraper
if "%secim%"=="5" goto excel_import
if "%secim%"=="6" goto project_check
if "%secim%"=="7" goto excel_export
goto exit

:ui
python iddaa_tahmin_ui.py
if errorlevel 1 (
    echo.
    echo  [HATA] Uygulama baslatilamadi. Hata ayrintisi yukarida yaziyor.
    echo  Eksik kutuphane bildirimi goruyorsan sunu calistirin:
    echo  pip install -r requirements.txt
    pause
)
goto exit

:today_scraper
python iddaa_oynanmamis_scraper.py
echo.
echo  Tamamlandi. Devam etmek icin bir tusa basin...
pause
goto exit

:scraper
echo.
set /p tarih1=  Baslangic tarihi (YYYY-MM-DD, orn: 2026-01-01): 
set /p tarih2=  Bitis tarihi    (YYYY-MM-DD, orn: 2026-03-31): 
python iddaa_scraper.py %tarih1% %tarih2%
echo.
echo  Tamamlandi. Devam etmek icin bir tusa basin...
pause
goto exit

:yearly_scraper
call tools\2026_veri_topla.bat
goto exit

:excel_import
call tools\excel_ice_aktar.bat
goto exit

:project_check
call tools\projeyi_kontrol_et.bat
goto exit

:excel_export
call tools\excel_disa_aktar.bat
goto exit
 
:exit
