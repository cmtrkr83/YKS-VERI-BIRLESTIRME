import os
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import *
from tkinter.messagebox import *
from tkinter import ttk

root = tk.Tk()
root.iconbitmap("odm.ico")
root.geometry("250x340")
root.resizable(False, False)
root.title("KODM  | YKS VERİ TOPLAMA")


def open_directory():
    global ana_klasor
    # Ana klasör yolunu belirtin
    ana_klasor = filedialog.askdirectory()


# DataFrame'leri depolamak için boş bir sözlük oluşturun
dataframeler = {}


def birlestir():
    # Alt klasörleri dolaşın
    for klasor, alt_klasorler, dosyalar in os.walk(ana_klasor):
        # Her Excel dosyası için döngü oluşturun
        for dosya in dosyalar:
            if dosya.endswith(
                ".xls"
            ):  # Sadece Excel dosyalarını işleyin, isterseniz .xls veya .csv olarak değiştirebilirsiniz
                dosya_yolu = os.path.join(klasor, dosya)

                # Excel dosyasını okuyun
                excel_verileri = pd.read_excel(dosya_yolu, sheet_name=None)

                # Her sayfayı ayrı ayrı DataFrame'e dönüştürerek kaydedin
                for sayfa_adi, veri in excel_verileri.items():
                    df = pd.DataFrame(veri)
                    #########################################################################
                    baslik = (str((df.iloc[1, 5]))).split("\n")
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

    for df in dataframeler:
        # Tümü boş olan sütunları silme işlemi
        dataframeler[df] = dataframeler[df].loc[
            :, dataframeler[df].notnull().any(axis=0)
        ]
    for sayfa_adi, df in dataframeler.items():
        dosya_adi = f"{sayfa_adi}_birlesmis.xlsx"
        df.to_excel(dosya_adi, index=False)


ttk.Label(root, text="YKS SONUÇLARINI BİRLEŞTİR", style="o.TLabel").pack(pady=5, padx=5)
ttk.Button(root, text="Gözat", command=open_directory, style="Custom.TButton").pack(
    pady=5
)
ttk.Button(root, text="Birleştir", command=birlestir, style="ask.TButton").pack(pady=5)
root.mainloop()