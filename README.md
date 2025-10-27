<!-- Badges -->
![Status](https://img.shields.io/badge/status-active-brightgreen.svg) ![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg) ![Platform](https://img.shields.io/badge/platform-macOS%20|%20Linux%20|%20Windows-lightgrey.svg)

# 🧩 YKS VERİ BİRLEŞTİRME

Bu proje, ÖSYM'den alınmış okul düzeyindeki YKS sonuçlarını içeren Excel (.xls) dosyalarını tarayıp, sayfa türlerine göre birleştirerek yıllık Excel çıktı dosyaları oluşturur.

Özet
- 🔎 GUI üzerinden bir ana klasör seçin (ana klasör: ilçe klasörlerini içeren dizin).
- 🗂️ Program, alt klasörlerdeki her `.xls` dosyasını okuyup, sayfa adına göre aynı yapıda olanları birleştirir.
- 📁 Oluşan birleştirilmiş Excel dosyaları `~/Documents` (Belgeler) klasörüne `YIL_<SayfaAdi>_birlesmis.xlsx` formatında kaydedilir.

Öne çıkan özellikler
- 🖥️ Basit Tkinter GUI (pencere ile çalışır).
- 📄 Desteklenen sayfa tipleri (kod içinde tanımlı isimlere göre):
	- ✅ `OkulAytTestSonucListesi`
	- ✅ `OkulYksTestSonucListesi`
	- ✅ `OkulYksPuanlariListesi`
	- ✅ `OkulYksYerlestirmePuanlariListe`
- ⚠️ Hatalı veya beklenmeyen formatta olan dosyalar atlanır; işlem sonunda hangi dosyaların atlandığı gösterilir.

Gereksinimler
- 🐍 Python 3.8 veya üzeri
- 📦 Paketler (pip ile kurulabilir): `pandas`, `xlrd`, `openpyxl`

Örnek kurulum

```bash
python3 -m pip install --user pandas xlrd openpyxl
```

Giriş — beklenen klasör yapısı

```text
input/          # ana giriş dizini (örnek)
	ilçe-1/
		okul-1/
			*.xls     # Okula ait OSYM çıktıları
		okul-2/
			*.xls
	ilçe-2/
		okul-1/
			*.xls
```

Notlar ve ipuçları
- ⚙️ Program şu an yalnızca `.xls` uzantısını tarar. Eğer `.xlsx` dosyalarınız varsa, `app.py` içinde uzantı kontrolünü `dosya.endswith(('.xls', '.xlsx'))` şeklinde genişletin ve gerekli engine'leri (`openpyxl`) kullanın.
- 🏷️ Kod, her Excel dosyasında belirli bir başlık yapısına (`df.iloc[1,5]`) dayanarak okul/ilçe bilgisini çıkarmaya çalışır; bu başlık formatı farklıysa dosya atlanır.
- 📣 Atlanan dosyalar işlem sonunda bir uyarıda listelenir; listeyi kontrol edip gerekiyorsa dosya formatını düzeltin.

Nasıl çalıştırılır

1. Gerekli paketleri kurun (yukarıya bakın).
2. Proje dizininde `app.py` dosyasını çalıştırın:

```bash
python3 app.py
```

3. Pencerede "Gözat" ile ana klasörü seçin → "Birleştir"e tıklayın.
4. İşlem tamamlandığında oluşan dosyalar `~/Documents` dizinine yazılır.

Çıktı örneği
- `2025_OkulYksTestSonucListesi_birlesmis.xlsx` gibi yıllık-sayfa bazlı dosyalar oluşur (yıl dinamik olarak belirlenir).

Ek: PyInstaller ile dağıtım

Proje kökünde `gui.spec` var; tek dosya uygulama yapmak isterseniz örnek:

```bash
pyinstaller --noconsole gui.spec
```

Sorun giderme
- ❗ Eğer `ImportError` veya "Dosya okunamadı" hatası alırsanız önce `pandas` ve `xlrd` kurulumunu kontrol edin.
- 🔍 Dosyalar atlanıyorsa, örnek atılan dosyadan başlık satırını (`df.iloc[1,5]`) inceleyin; kod okul/ilçe bilgisini buradan çıkarır.

Katkıda bulunmak isterseniz
- ✍️ `app.py` içinde küçük geliştirmeler yapabilirsiniz. Öneriler:
	- `.xlsx` desteği ekleyin
	- Hata/uyarı log'unu bir dosyaya yazdırın

-----------------------------------

2025 — Güncel, sade ve görselleştirilmiş README
