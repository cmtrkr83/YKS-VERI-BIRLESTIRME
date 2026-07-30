import os, shutil, datetime, warnings, logging
from tkinter import *
from tkinter import messagebox
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
import subprocess

warnings.filterwarnings("ignore", category=UserWarning, module="xlrd")
logging.getLogger("xlrd").setLevel(logging.ERROR)

GECERLI_SAYFALAR = [
    "OkulAytTestSonucListesi",
    "OkulYksTestSonucListesi",
    "OkulYksPuanlariListesi",
    "OkulYksYerlestirmePuanlariListe"
]

log_satirlari = []
def log_ekle(seviye, mesaj):
    z = datetime.datetime.now().strftime("%H:%M:%S")
    s = f"[{z}] [{seviye}] {mesaj}"
    log_satirlari.append(s)
    print(mesaj)

def log_kaydet(dosyayolu):
    with open(dosyayolu, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("  YKS VERI BIRLESTIRME LOG\n")
        f.write(f"  Tarih: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        for satir in log_satirlari:
            f.write(satir + "\n")

def hata_kopyala(hata_tipi, dosya_yolu, hata_ana_klasor):
    alt = os.path.join(hata_ana_klasor, hata_tipi)
    os.makedirs(alt, exist_ok=True)
    hedef = os.path.join(alt, os.path.basename(dosya_yolu))
    if os.path.exists(hedef):
        ad, uz = os.path.splitext(os.path.basename(dosya_yolu))
        hedef = os.path.join(alt, f"{ad}_{datetime.datetime.now().strftime('%H%M%S')}{uz}")
    shutil.copy2(dosya_yolu, hedef)

def ilce_okul_bul(dosya_yolu):
    try:
        b = dosya_yolu.replace("\\", "/").split("/")
        if len(b) >= 3:
            ilce = b[-3].replace("-YKS", "").replace(" YKS", "").strip()
            return ilce, b[-2].strip()
    except:
        pass
    return "", os.path.basename(os.path.dirname(dosya_yolu))

root = tk.Tk()
root.title("KODM | YKS VERI TOPLAMA")
root.geometry("400x180")
root.resizable(False, False)
output_folder = os.path.join(os.path.expanduser("~"), "Documents", "yks verileri")

def durum_guncelle(mesaj):
    """Durum etiketini guncelle ve arayuzu yenile"""
    durum_label.config(text=mesaj)
    root.update()
    log_ekle("DURUM", mesaj)

def open_directory():
    global ana_klasor
    ana_klasor = filedialog.askdirectory()
    selected_folder_label.config(text=ana_klasor if ana_klasor else "Klasor secilmedi.")

def birlestir():
    global log_satirlari
    log_satirlari = []
    dataframeler = {}
    
    log_ekle("BASLA", f"Klasor: {ana_klasor}")
    hata_ana_klasor = os.path.join(output_folder, "HATA_ALINAN_DOSYALAR")
    if os.path.exists(hata_ana_klasor): shutil.rmtree(hata_ana_klasor)
    os.makedirs(hata_ana_klasor)

    # 1. ADIM: Tum dosyalari tara, okul bazinda grupla
    okul_dosyalari = {}  # (ilce, okul) -> [dosya_yolu, ...]
    tum_dosyalar = []
    
    for klasor, _, dosyalar in os.walk(ana_klasor):
        for dosya in dosyalar:
            if dosya.endswith((".xls", ".xlsx")):
                fp = os.path.join(klasor, dosya)
                ilce, okul = ilce_okul_bul(fp)
                tum_dosyalar.append(fp)
                anahtar = (ilce, okul)
                if anahtar not in okul_dosyalari:
                    okul_dosyalari[anahtar] = []
                okul_dosyalari[anahtar].append(fp)
    
    durum_guncelle(f"{len(tum_dosyalar)} dosya, {len(okul_dosyalari)} okul bulundu. Analiz basliyor...")
    log_ekle("BASLA", f"{len(tum_dosyalar)} dosya, {len(okul_dosyalari)} okul bulundu.")
    progress["maximum"] = len(tum_dosyalar)
    progress["value"] = 0

    # 2. ADIM: Her okul icin hangi sayfalari gondermis analizi
    okul_raporlari = {}
    basari_sayisi = 0
    islenen = 0

    for anahtar, dosyalar in okul_dosyalari.items():
        ilce, okul = anahtar
        okul_sayfalari = {s: [] for s in GECERLI_SAYFALAR}  # sayfa -> [dosya_yollari]
        okul_hatali = []  # hic okunamayan dosyalar
        okul_baslik_hatali = []  # baslik bulunamayan
        
        for fp in dosyalar:
            try:
                excel = pd.read_excel(fp, sheet_name=None)
                dosyadaki_sayfalar = list(excel.keys())
                bulunan = [s for s in dosyadaki_sayfalar if s in GECERLI_SAYFALAR]
                
                if not bulunan:
                    okul_hatali.append((fp, f"Gecerli sayfa bulunamadi (mevcut: {dosyadaki_sayfalar})"))
                    hata_kopyala("SAYFA_BULUNAMADI", fp, hata_ana_klasor)
                    islenen += 1; progress["value"] = islenen; root.update_idletasks()
                    continue
                
                for sayfa in bulunan:
                    df = pd.DataFrame(excel[sayfa])
                    try:
                        baslik = str(df.iloc[1, 5]).split("\n")[1].split("(")
                        okul_adi = baslik[0].strip()
                        ilce_adi = baslik[1].split()[2][:-1]
                        df.insert(2, "ILCE", ""); df.insert(3, "OKUL", "")
                        df["OKUL"] = okul_adi; df["ILCE"] = ilce_adi
                    except:
                        okul_baslik_hatali.append(fp)
                        hata_kopyala("BASLIK_BULUNAMADI", fp, hata_ana_klasor)
                        islenen += 1; progress["value"] = islenen; root.update_idletasks()
                        continue

                    # Sutun sayisina gore icerik dogrulama (basit)
                    sutun_sayisi = len(df.columns)
                    if sayfa == "OkulAytTestSonucListesi":
                        okul_sayfalari[sayfa].append({"fp": fp, "df": df, "sutun": sutun_sayisi, "uyari": ""})
                    elif sayfa == "OkulYksTestSonucListesi":
                        okul_sayfalari[sayfa].append({"fp": fp, "df": df, "sutun": sutun_sayisi, "uyari": ""})
                    elif sayfa == "OkulYksPuanlariListesi":
                        okul_sayfalari[sayfa].append({"fp": fp, "df": df, "sutun": sutun_sayisi, "uyari": ""})
                    elif sayfa == "OkulYksYerlestirmePuanlariListe":
                        okul_sayfalari[sayfa].append({"fp": fp, "df": df, "sutun": sutun_sayisi, "uyari": ""})
                    
                    basari_sayisi += 1
                    
            except Exception as hata:
                okul_hatali.append((fp, f"Dosya okunamadi: {str(hata)[:80]}"))
                hata_kopyala("DOSYA_OKUNAMADI", fp, hata_ana_klasor)
            
            islenen += 1
            progress["value"] = islenen
            root.update_idletasks()
        
        okul_raporlari[anahtar] = {
            "dosyalar": okul_sayfalari,
            "hatali": okul_hatali,
            "baslik_hatali": okul_baslik_hatali
        }

    # 3. ADIM: Rapor olustur
    rapor_satirlari = []
    rapor_satirlari.append("=" * 70)
    rapor_satirlari.append("  OKUL BAZINDA SAYFA DURUM RAPORU")
    rapor_satirlari.append(f"  Tarih: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    rapor_satirlari.append("=" * 70)
    rapor_satirlari.append(f"\nToplam okul: {len(okul_raporlari)}")
    rapor_satirlari.append(f"Toplam dosya: {len(tum_dosyalar)}\n")
    rapor_satirlari.append("-" * 70)

    eksikli_okul_sayisi = 0
    toplam_eksik = {s: 0 for s in GECERLI_SAYFALAR}
    toplam_fazla = {s: 0 for s in GECERLI_SAYFALAR}
    
    for anahtar in sorted(okul_raporlari.keys(), key=lambda x: (x[0], x[1])):
        ilce, okul = anahtar
        info = okul_raporlari[anahtar]
        okul_sorunlu = False
        okul_satirlari = []
        
        okul_satirlari.append(f"\n[{ilce}] {okul}")
        
        for sayfa in GECERLI_SAYFALAR:
            dosya_listesi = info["dosyalar"][sayfa]
            sayfa_adi_kisa = sayfa.replace("Okul", "").replace("Listesi", "").replace("Liste", "")
            
            if len(dosya_listesi) == 0:
                okul_satirlari.append(f"  \u2716 {sayfa_adi_kisa} -> EKSIK (dosya gonderilmemis)")
                toplam_eksik[sayfa] += 1
                okul_sorunlu = True
            elif len(dosya_listesi) == 1:
                detay = dosya_listesi[0]
                uyari = detay["uyari"]
                if uyari:
                    okul_satirlari.append(f"  \u26a0 {sayfa_adi_kisa} -> VAR (uyari: {uyari})")
                    okul_sorunlu = True
                else:
                    okul_satirlari.append(f"  \u2714 {sayfa_adi_kisa} -> VAR")
            else:
                okul_satirlari.append(f"  \u26a0 {sayfa_adi_kisa} -> {len(dosya_listesi)} KEZ (fazla dosya)")
                toplam_fazla[sayfa] += (len(dosya_listesi) - 1)
                okul_sorunlu = True
        
        if info["hatali"]:
            for fp, sebep in info["hatali"]:
                dosya_adi = os.path.basename(fp)
                okul_satirlari.append(f"  \u2716 HATA: {dosya_adi} -> {sebep}")
            okul_sorunlu = True
        
        if info["baslik_hatali"]:
            for fp in info["baslik_hatali"]:
                okul_satirlari.append(f"  \u2716 BASLIK HATASI: {os.path.basename(fp)}")
            okul_sorunlu = True
        
        if okul_sorunlu:
            eksikli_okul_sayisi += 1
            for satir in okul_satirlari:
                rapor_satirlari.append(satir)
                log_ekle("OKUL", satir)
        # Sorunsuz okullar log'a eklenmez

    # OZET
    rapor_satirlari.append("\n" + "=" * 70)
    rapor_satirlari.append("OZET TABLO")
    rapor_satirlari.append("=" * 70)
    rapor_satirlari.append(f"Toplam okul: {len(okul_raporlari)}")
    rapor_satirlari.append(f"Sorunlu okul: {eksikli_okul_sayisi}")
    rapor_satirlari.append(f"Sorunsuz okul: {len(okul_raporlari) - eksikli_okul_sayisi}")
    rapor_satirlari.append("")
    for sayfa in GECERLI_SAYFALAR:
        kisa = sayfa.replace("Okul", "").replace("Listesi", "").replace("Liste", "")
        rapor_satirlari.append(f"  {kisa}: Eksik={toplam_eksik[sayfa]}, Fazla={toplam_fazla[sayfa]}")

    # 4. ADIM: Verileri gercek birlestirme ve cikti
    durum_guncelle("Excel dosyalari birlestiriliyor...")
    sayac = 0
    toplam_veri = sum(len(info["dosyalar"][s]) for info in okul_raporlari.values() for s in GECERLI_SAYFALAR)
    for anahtar, info in okul_raporlari.items():
        for sayfa in GECERLI_SAYFALAR:
            for dosya_info in info["dosyalar"][sayfa]:
                df = dosya_info["df"]
                if sayfa not in dataframeler:
                    dataframeler[sayfa] = df
                else:
                    dataframeler[sayfa] = pd.concat([dataframeler[sayfa], df])
                sayac += 1
                if sayac % 50 == 0:
                    durum_guncelle(f"Excel birlestiriliyor... {sayac}/{toplam_veri}")

    durum_guncelle("Bos sutunlar temizleniyor...")
    root.update()
    for df_adi in dataframeler:
        dataframeler[df_adi] = dataframeler[df_adi].loc[:, dataframeler[df_adi].notnull().any(axis=0)]

    durum_guncelle("Excel ciktilari yaziliyor...")
    for sayfa_adi, df in dataframeler.items():
        yil = pd.to_datetime("today").year
        dosya_adi = os.path.join(output_folder, f"{yil}_{sayfa_adi}_birlesmis.xlsx")
        durum_guncelle(f"Kaydediliyor: {sayfa_adi[:20]}... ({len(df)} satir)")
        df.to_excel(dosya_adi, index=False)
        log_ekle("CIKTI", f"{len(df)} satir -> {dosya_adi}")

    durum_guncelle("Rapor hazirlaniyor...")
    zaman_damgasi = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    rapor_dosyasi = os.path.join(output_folder, f"OKUL_RAPORU_{zaman_damgasi}.txt")
    with open(rapor_dosyasi, "w", encoding="utf-8") as f:
        f.write("\n".join(rapor_satirlari))
    
    log_dosyasi = os.path.join(output_folder, f"YKS_LOG_{zaman_damgasi}.txt")
    log_kaydet(log_dosyasi)

    durum_guncelle("Islem tamamlandi!")
    messagebox.showinfo(
        "Islem Tamamlandi",
        f"\u2705 Islem basariyla tamamlandi!\n\n"
        f"Basarili: {basari_sayisi}\n"
        f"Sorunlu okul: {eksikli_okul_sayisi}\n"
        f"Sorunsuz okul: {len(okul_raporlari) - eksikli_okul_sayisi}\n\n"
        f"Detayli rapor:\n{rapor_dosyasi}\n"
        f"Log:\n{log_dosyasi}"
    )
    durum_guncelle("Hazir. Yeni bir klasor secip tekrar baslatabilirsiniz.")

def open_output_folder():
    subprocess.Popen(f'explorer "{output_folder}"')

ttk.Label(root, text="YKS SONUCLARINI BIRLESTIR", font=("Arial", 10, "bold")).pack(pady=5)
selected_folder_label = ttk.Label(root, text="Klasor secilmedi.", wraplength=300)
selected_folder_label.pack(pady=2)
button_frame = ttk.Frame(root)
button_frame.pack(pady=5)
ttk.Button(button_frame, text="Gozat", command=open_directory).pack(side=LEFT, padx=5)
ttk.Button(button_frame, text="Birlestir", command=birlestir).pack(side=LEFT, padx=5)
ttk.Button(button_frame, text="Klasoru Ac", command=open_output_folder).pack(side=LEFT, padx=5)
progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress.pack(pady=5, padx=5, fill="x", expand=True)
durum_label = ttk.Label(root, text="", wraplength=380, foreground="#555")
durum_label.pack(pady=2)
root.mainloop()
