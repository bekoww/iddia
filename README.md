# Iddaa Oran Eslestirici

Python tabanli masaustu tahmin ve veri toplama araci.

## Baslatma

Ana menu:

```bat
baslat.bat
```

Gunluk kullanim icin tek giris dosyasi budur.

## Veri Toplama

2026 verisini kaldigi yerden toplamak icin:

`baslat.bat` icinden `2026 Verisini Kolay Topla` sec.

Komut satirindan:

```powershell
python iddaa_scraper.py 2026-01-01 2026-09-08 -o data/iddaa.sqlite --workers 12
```

Yarida kalirsa ayni komutu tekrar calistir. Mevcut Excel okunur, kayitli maclar atlanir.

Sifirdan baslatmak icin:

```powershell
python iddaa_scraper.py 2026-01-01 2026-09-08 -o data/iddaa.sqlite --workers 12 --fresh
```

## Excel Iceri Aktarma

Gunluk, aylik veya toplu Excel dosyalarini ana veriye eklemek icin:

`baslat.bat` icinden `Excel Verilerini Ice Aktar` sec.

Bu arac:

- Ana dosya olarak `excel/bahis_oranlari_2026.xlsx` kullanir ve bu dosya icin yedek alir.
- `Tarih + Saat + Lig + Mac` ayniysa tekrar eden maclari temizler.
- Son gelen kaydi tutar.
- Kac yeni mac ve kac tekrar/guncelleme oldugunu raporlar.

## Ana Veri Deposu

- Scraper ve uygulama ana veri olarak `data/iddaa.sqlite` kullanir.
- `excel` klasorundeki dosyalar arsiv, ice aktarma ve raporlama icindir.
- `Excel Verilerini Ice Aktar` menusu, dosyalari birlestirdikten sonra SQLite
  deposunu da yeniler.
- Guncel raporu almak icin `tools\\excel_disa_aktar.bat` calistirilir; dosya
  `excel/bahis_oranlari_guncel.xlsx` olarak olusur.

## Uygulama Icindeki Yeni Ozellikler

- Veri kapsami: `Tum Veri`, `Sadece 2026`, `Son 365 Gun`, `Son 180 Gun`
- Kupon risk modu: `Guvenli`, `Dengeli`, `Agresif`
- Kupon toplam oran hesabi
- Veri Sagligi sekmesi
- Veriyi Guncelle butonu: ana verideki son tarihten dune kadar eksik gecmis veriyi ceker.
- Tahmin gecmisi durum takibi: `Bekliyor`, `Tuttu`, `Yatti`

## Test

```bat
baslat.bat
```

Menuden `Projeyi Kontrol Et` sec.

veya:

```powershell
python -m compileall -q .
python -m unittest discover -s tests -v
```
