@echo off
chcp 65001 >nul
title İddaa Oran Eşleştirici - Tahmin Arayüzü
cd /d "%~dp0"
python iddaa_tahmin_ui.py
if errorlevel 1 (
    echo.
    echo  [HATA] Uygulama baslatilamadi. Hata ayrintisi yukarida yaziyor.
    echo  Eksik kutuphane bildirimi goruyorsan sunu calistirin:
    echo  pip install -r requirements.txt
    pause
)
