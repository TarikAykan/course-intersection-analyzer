@echo off
setlocal

echo DersKesisimAnalizi icin exe olusturma basladi...

if exist assets\app_icon.ico (
    echo Ikon bulundu: assets\app_icon.ico
    pyinstaller --onefile --windowed --icon=assets/app_icon.ico --name DersKesisimAnalizi main.py
) else (
    echo Ikon bulunamadi, ikon olmadan devam ediliyor.
    pyinstaller --onefile --windowed --name DersKesisimAnalizi main.py
)

echo.
echo Islem tamamlandi. Cikti dosyasini dist klasorunde bulabilirsiniz.
pause
