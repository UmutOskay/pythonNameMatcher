# Excel Name Matcher

Excel dosyalarındaki isimleri eşleştiren GUI uygulaması.

## Özellikler

- ✅ İki Excel dosyası arasında isim eşleştirme
- ✅ Türkçe karakter desteği (İ, ı, Ş, Ğ, Ü, Ö, Ç)
- ✅ Kullanıcı dostu GUI
- ✅ Excel sütunlarını harf ile seçim (A, B, C, AA, AB...)
- ✅ Detaylı raporlama

## Kullanım

### Windows EXE (Python kurulumu gerektirmez)

1. [Releases](../../releases) sayfasından `ExcelNameMatcher.exe` dosyasını indirin
2. Çift tıklayarak çalıştırın
3. Excel dosyalarını seçin
4. Sütunları belirtin (örn: A, B, E)
5. "Eşleştirmeyi Başlat" butonuna tıklayın

### Python ile Çalıştırma

```bash
pip install pandas openpyxl
python excel_matcher_gui.py
```

## Geliştirme

Bu proje GitHub Actions ile otomatik olarak Windows EXE oluşturur.

## Lisans

Özel kullanım için geliştirilmiştir.

