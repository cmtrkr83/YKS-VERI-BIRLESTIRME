import os
from tkinter import *
from tkinter import messagebox
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
import subprocess



root = tk.Tk()
root.title("KODM  | YKS VERİ TOPLAMA")
root.geometry("350x150")
root.resizable(False, False)

output_folder = os.path.join(os.path.expanduser("~"), "Documents")

def open_directory():
    global ana_klasor
    ana_klasor = filedialog.askdirectory()
    selected_folder_label.config(text=ana_klasor if ana_klasor else "Klasör seçilmedi.")

dataframeler = {}


def birlestir():
    dataframeler.clear()
    # Tüm işlenecek dosyaları önceden say
    toplam_dosya = 0
    for klasor, alt_klasorler, dosyalar in os.walk(ana_klasor):
        for dosya in dosyalar:
            if dosya.endswith(".xls"):
                toplam_dosya += 1

    progress["maximum"] = toplam_dosya
    progress["value"] = 0
    root.update_idletasks()

    atlanan_dosyalar = []

    islenen = 0
    # Alt klasörleri dolaşın
    for klasor, alt_klasorler, dosyalar in os.walk(ana_klasor):
        # Her Excel dosyası için döngü oluşturun
        for dosya in dosyalar:
            if dosya.endswith(
                ".xls"
            ):  # Sadece Excel dosyalarını işleyin, isterseniz .xls veya .csv olarak değiştirebilirsiniz
                dosya_yolu = os.path.join(klasor, dosya)
                try:
                # Excel dosyasını okuyun
                    excel_verileri = pd.read_excel(dosya_yolu, sheet_name=None)

                # Her sayfayı ayrı ayrı DataFrame'e dönüştürerek kaydedin
                    for sayfa_adi, veri in excel_verileri.items():
                        df = pd.DataFrame(veri)
                    #########################################################################
                        try:
                            baslik = (str((df.iloc[1, 5]))).split("\n")
                            print(dosya_yolu)
                            baslik = baslik[1]
                            baslik = baslik.split("(")
                            okul = baslik[0]
                            ilce = baslik[1]
                            ilce = ilce.split()
                            ilce = ilce[2]
                            ilce = ilce[: len(ilce) - 1]
                            okul = okul.strip()
                            df.insert(2, "İLÇE", "")
                            df.insert(3, "OKUL", "")
                            df["OKUL"] = okul
                            df["İLÇE"] = ilce
                        except Exception:
                            print(f"Başlık bulunamadı: {dosya_yolu}")
                            atlanan_dosyalar.append(dosya_yolu)
                            continue
                    ############################################################################
                    if sayfa_adi == "OkulAytTestSonucListesi":
                        satir_atlama_miktari = 11
                        df = df.iloc[satir_atlama_miktari:]  # Satır atlama işlemi
                        df = df.rename(
                            columns={
                                "Unnamed: 1": "TC",
                                "Unnamed: 4": "İSİM",
                                "Unnamed: 6": "TDE-DOĞRU",
                                "Unnamed: 7": "TDE-YANLIŞ",
                                "Unnamed: 8": "TDE-BOŞ",
                                "Unnamed: 9": "TDE-HAM",
                                "Unnamed: 10": "TARİH(1)-DOĞRU",
                                "Unnamed: 11": "TARİH(1)-YANLIŞ",
                                "Unnamed: 12": "TARİH(1)-BOŞ",
                                "Unnamed: 13": "TARİH(1)-HAM",
                                "Unnamed: 14": "COĞRAFYA(1)-DOĞRU",
                                "Unnamed: 15": "COĞRAFYA(1)-YANLIŞ",
                                "Unnamed: 16": "COĞRAFYA(1)-BOŞ",
                                "Unnamed: 17": "COĞRAFYA(1)-HAM",
                                "Unnamed: 18": "TARİH(2)-DOĞRU",
                                "Unnamed: 19": "TARİH(2)-YANLIŞ",
                                "Unnamed: 20": "TARİH(2)-BOŞ",
                                "Unnamed: 21": "TARİH(2)-HAM",
                                "Unnamed: 22": "COĞRAFYA(2)-DOĞRU",
                                "Unnamed: 23": "COĞRAFYA(2)-YANLIŞ",
                                "Unnamed: 24": "COĞRAFYA(2)-BOŞ",
                                "Unnamed: 25": "COĞRAFYA(2)-HAM",
                                "Unnamed: 26": "FELSEFE-DOĞRU",
                                "Unnamed: 27": "FELSEFE-YANLIŞ",
                                "Unnamed: 28": "FELSEFE-BOŞ",
                                "Unnamed: 29": "FELSEFE-HAM",
                                "Unnamed: 30": "DİNKÜLT-DOĞRU",
                                "Unnamed: 31": "DİNKÜLT-YANLIŞ",
                                "Unnamed: 32": "DİNKÜLT-BOŞ",
                                "Unnamed: 33": "DİNKÜLT-HAM",
                                "Unnamed: 35": "MATEMATİK-DOĞRU",
                                "Unnamed: 37": "MATEMATİK-YANLIŞ",
                                "Unnamed: 38": "MATEMATİK-BOŞ",
                                "Unnamed: 39": "MATEMATİK-HAM",
                                "Unnamed: 41": "FİZİK-DOĞRU",
                                "Unnamed: 42": "FİZİK-YANLIŞ",
                                "Unnamed: 43": "FİZİK-BOŞ",
                                "Unnamed: 44": "FİZİK-HAM",
                                "Unnamed: 45": "KİMYA-DOĞRU",
                                "Unnamed: 46": "KİMYA-YANLIŞ",
                                "Unnamed: 47": "KİMYA-BOŞ",
                                "Unnamed: 48": "KİMYA-HAM",
                                "Unnamed: 49": "BİYOLOJİ-DOĞRU",
                                "Unnamed: 50": "BİYOLOJİ-YANLIŞ",
                                "Unnamed: 51": "BİYOLOJİ-BOŞ",
                                "Unnamed: 52": "BİYOLOJİ-HAM",
                            }
                        )
                    elif sayfa_adi == "OkulYksTestSonucListesi":
                        satir_atlama_miktari = 11
                        df = df.iloc[satir_atlama_miktari:]  # Satır atlama işlemi
                        df = df.rename(
                            columns={
                                "Unnamed: 1": "TC",
                                "Unnamed: 4": "İSİM",
                                "Unnamed: 6": "TÜRKÇE(TYT)-DOĞRU",
                                "Unnamed: 7": "TÜRKÇE(TYT)-YANLIŞ",
                                "Unnamed: 8": "TÜRKÇE(TYT)-BOŞ",
                                "Unnamed: 9": "TÜRKÇE(TYT)-HAM",
                                "Unnamed: 10": "SOSBİL(TYT)-DOĞRU",
                                "Unnamed: 11": "SOSBİL(TYT)-YANLIŞ",
                                "Unnamed: 12": "SOSBİL(TYT)-BOŞ",
                                "Unnamed: 13": "SOSBİL(TYT)-HAM",
                                "Unnamed: 14": "TEMELMAT(TYT)-DOĞRU",
                                "Unnamed: 15": "TEMELMAT(TYT)-YANLIŞ",
                                "Unnamed: 16": "TEMELMAT(TYT)-BOŞ",
                                "Unnamed: 17": "TEMELMAT(TYT)-HAM",
                                "Unnamed: 18": "FENBİL(TYT)-DOĞRU",
                                "Unnamed: 19": "FENBİL(TYT)-YANLIŞ",
                                "Unnamed: 20": "FENBİL(TYT)-BOŞ",
                                "Unnamed: 21": "FENBİL(TYT)-HAM",
                                "Unnamed: 22": "TDESOS(AYT)-DOĞRU",
                                "Unnamed: 23": "TDESOS(AYT)-YANLIŞ",
                                "Unnamed: 24": "TDESOS(AYT)-BOŞ",
                                "Unnamed: 25": "TDESOS(AYT)-HAM",
                                "Unnamed: 28": "SOSBİL(AYT)-DOĞRU",
                                "Unnamed: 29": "SOSBİL(AYT)-YANLIŞ",
                                "Unnamed: 30": "SOSBİL(AYT)-BOŞ",
                                "Unnamed: 31": "SOSBİL(AYT)-HAM",
                                "Unnamed: 32": "MATEMATİK(AYT)-DOĞRU",
                                "Unnamed: 33": "MATEMATİK(AYT)-YANLIŞ",
                                "Unnamed: 34": "MATEMATİK(AYT)-BOŞ",
                                "Unnamed: 35": "MATEMATİK(AYT)-HAM",
                                "Unnamed: 36": "FENBİL(AYT)-DOĞRU",
                                "Unnamed: 37": "FENBİL(AYT)-YANLIŞ",
                                "Unnamed: 38": "FENBİL(AYT)-BOŞ",
                                "Unnamed: 39": "FENBİL(AYT)-HAM",
                                "Unnamed: 40": "YABDİL(YDT)-DOĞRU",
                                "Unnamed: 41": "YABDİL(YDT)-YANLIŞ",
                                "Unnamed: 42": "YABDİL(YDT)-BOŞ",
                                "Unnamed: 43": "YABDİL(YDT)-HAM",
                            }
                        )
                    elif sayfa_adi == "OkulYksPuanlariListesi":
                        satir_atlama_miktari = 10
                        df = df.iloc[satir_atlama_miktari:]  # Satır atlama işlemi
                        df = df.rename(
                            columns={
                                "Unnamed: 1": "TC",
                                "Unnamed: 3": "İSİM",
                                "Unnamed: 6": "TYT PUAN",
                                "Unnamed: 7": "TYT BAŞARI SIRASI",
                                "Unnamed: 8": "SAYISAL PUAN",
                                "Unnamed: 9": "SAYISAL BAŞARI SIRASI",
                                "Unnamed: 10": "SÖZEL PUAN",
                                "Unnamed: 11": "SÖZEL BAŞARI SIRASI",
                                "Unnamed: 12": "EA PUAN",
                                "Unnamed: 13": "EA BAŞARI SIRASI",
                                "Unnamed: 14": "YABANCI DİL PUAN",
                                "Unnamed: 15": "YABANCI DİL BAŞARI SIRASI ",
                            }
                        )
                    elif sayfa_adi == "OkulYksYerlestirmePuanlariListe":
                        satir_atlama_miktari = 10
                        df = df.iloc[satir_atlama_miktari:]  # Satır atlama işlemi
                        df = df.rename(
                            columns={
                                "Unnamed: 1": "TC",
                                "Unnamed: 3": "İSİM",
                                "Unnamed: 6": "TYT PUAN",
                                "Unnamed: 7": "TYT BAŞARI SIRASI",
                                "Unnamed: 8": "SAYISAL PUAN",
                                "Unnamed: 9": "SAYISAL BAŞARI SIRASI",
                                "Unnamed: 10": "SÖZEL PUAN",
                                "Unnamed: 11": "SÖZEL BAŞARI SIRASI",
                                "Unnamed: 12": "EA PUAN",
                                "Unnamed: 13": "EA BAŞARI SIRASI",
                                "Unnamed: 14": "YABANCI DİL PUAN",
                                "Unnamed: 15": "YABANCI DİL BAŞARI SIRASI",
                            }
                        )
                    else:
                        print("Sayfa İsimleri Uygun Değil")
                    if sayfa_adi not in dataframeler:
                        dataframeler[sayfa_adi] = df
                    else:
                        dataframeler[sayfa_adi] = pd.concat(
                            [dataframeler[sayfa_adi], df]
                        )
                except Exception:
                    print(f"Dosya okunamadı: {dosya_yolu}")
                    atlanan_dosyalar.append(dosya_yolu)
                islenen += 1
                progress["value"] = islenen
                root.update_idletasks()
    for df in dataframeler:
        # Tümü boş olan sütunları silme işlemi
        dataframeler[df] = dataframeler[df].loc[
            :, dataframeler[df].notnull().any(axis=0)
        ]
    for sayfa_adi, df in dataframeler.items():
        yil = pd.to_datetime("today").year
        dosya_adi = os.path.join(output_folder, f"{yil}_{sayfa_adi}_birlesmis.xlsx")
        df.to_excel(dosya_adi, index=False)
    if atlanan_dosyalar:
        messagebox.showwarning(
            "Atlanan Dosyalar",
            "Bazı dosyalar işlenemedi:\n" + "\n".join(atlanan_dosyalar)
        )
    messagebox.showinfo(
        "İşlem Tamamlandı",
        "İşlem tamamlandı! Birleşmiş dosyalar Belgeler dizininde oluşturuldu.",
    )
def open_output_folder():
    output_folder = os.path.join(os.path.expanduser("~"), "Documents")
    subprocess.Popen(f'explorer "{output_folder}"')

ttk.Label(root, text="YKS SONUÇLARINI BİRLEŞTİR", style="o.TLabel").pack(pady=5, padx=5)
selected_folder_label = ttk.Label(root, text="Klasör seçilmedi.", wraplength=220)
selected_folder_label.pack(pady=2)
button_frame = ttk.Frame(root)
button_frame.pack(pady=5)

ttk.Button(button_frame, text="Gözat", command=open_directory, style="Custom.TButton").pack(side=LEFT, padx=5)
ttk.Button(button_frame, text="Birleştir", command=birlestir, style="ask.TButton").pack(side=LEFT, padx=5)
ttk.Button(button_frame, text="Klasörü Aç", command=open_output_folder).pack(side=LEFT, padx=5)
progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress.pack(pady=10, padx=5, fill="x", expand=True)


    
root.mainloop()