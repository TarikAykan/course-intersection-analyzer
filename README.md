# Ders Kesisim Analizi

Bu masaustu uygulamasi, Excel veya API kaynakli ogrenci/ders kayitlarini analiz ederek iki secili dersi ayni anda alan ogrencileri listeler.

## Program ne ise yarar?

- Kullanici veri kaynagi olarak Excel veya API secer.
- Sutunlar eslestirilir: Ogrenci No (opsiyonel), Ad, Soyad, Ders Adi.
- Iki ders secilir (ornek: Enerji ve Afet).
- Program bu iki dersi de alan ogrencileri bulur.
- Sonuclar tabloda gosterilir ve istenirse Excel'e aktarilir.

## Kurulum

1. Python 3.10+ kurulu olmali.
2. Proje klasorunde terminal acin.
3. Paketleri kurun:

```bash
pip install -r requirements.txt
```

## Uygulamayi calistirma

```bash
python main.py
```

## Kullanım Adimlari

1. **Veri Kaynagini Sec**
   - Excel icin `Veri Kaynagi: Excel` secin ve `Excel Dosyasi Sec` butonuna basin.
   - API icin `Veri Kaynagi: API` secin, URL girin ve `API Verisini Cek` butonuna basin.
2. **Sutunlari Eslestir**
   - Ogrenci No, Ad, Soyad ve Ders Adi sutunlarini secin.
   - Eger ogrenci no yoksa `Ogrenci no yok` secenegini kullanin.
   - `Sutunlari Onayla` butonuna basin.
3. **Dersleri Sec**
   - Birinci ve ikinci dersi secin.
4. **Analiz Et**
   - `Analiz Et` butonuna basin.
   - Sonuclar tabloda gosterilir.
5. **Disa Aktar**
   - `Sonucu Excel'e Aktar` butonu ile sonucu kaydedin.

## Excel Veri Yapisi

Program farkli OBS ciktilariyla calisabilsin diye sutunlari kullanici secer.
Beklenen alanlar:
- `Ogrenci No` (opsiyonel)
- `Ad`
- `Soyad`
- `Ders Adi`

## Ozel Notlar

- Eslestirme onceligi **Ogrenci No** uzerindedir.
- Ogrenci No yoksa **Ad + Soyad** birlestirilerek eslestirme yapilir.
- Analiz yalnizca secilen **iki dersin kesisimini** hesaplar.
- Kod yapisi ileride ucuncu ders filtresi eklenebilecek sekilde modulerdir.

## EXE Olusturma (PyInstaller)

Windows'ta proje klasorunde su dosyayi calistirin:

```bat
build_exe.bat
```

Bu script su komutu kullanir:

```bash
pyinstaller --onefile --windowed --name DersKesisimAnalizi main.py
```

Eger `assets/app_icon.ico` varsa otomatik olarak ikon parametresi ile build alir.

## Proje Yapisi

```text
student_course_checker/
│
├── main.py
├── requirements.txt
├── README.md
├── build_exe.bat
│
├── app/
│   ├── __init__.py
│   ├── gui.py
│   ├── excel_reader.py
│   ├── analyzer.py
│   └── exporter.py
│
├── assets/
│   └── app_icon.ico
│
└── sample_data/
    ├── ornek_veri.csv
    ├── ornek_api_data.json
    └── create_sample_excel.py
```

## Ornek Test Verisi

- `sample_data/ornek_veri.csv`: Hazir ornek veri.
- `sample_data/ornek_api_data.json`: API simulasyonu icin JSON veri.
- `sample_data/create_sample_excel.py`: Bu veriyi `.xlsx` formatina donusturmek icin script.

Script calistirma:

```bash
python sample_data/create_sample_excel.py
```

Olusan dosya: `sample_data/ornek_veri.xlsx`
