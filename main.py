import PyPDF2
import re
import os
import pandas as pd
from tkinter import Tk, filedialog


def pdf_den_kodlari_al(pdf_yolu):
    kodlar = []

    with open(pdf_yolu, 'rb') as pdf_dosyasi:
        pdf_okuyucu = PyPDF.PdfFileReader(pdf_dosyasi)

        for sayfa_sayisi in range(pdf_okuyucu.getNumPages()):
            sayfa = pdf_okuyucu.getPage(sayfa_sayisi)
            metin = sayfa.extractText()

            # Parantez içindeki sadece ilk karakterin harf olup olmadığını kontrol et
            matches = find_codes_in_text(metin, sayfa_sayisi + 1)

            # Eğer bir eşleşme varsa, listeye ekle
            kodlar.extend(matches)

    return kodlar



def find_codes_in_text(text, sayfa_numarasi):
    kodlar = []
    lines = text.split('\n')

    for line in lines:
        # Parantez içindeki sadece ilk karakterin harf olup olmadığını kontrol et
        matches = re.findall(r'\((\b[A-Za-z]\b[^)]*)\)', line)

        # Eğer bir eşleşme varsa, listeye ekle
        if matches:
            kodlar.extend([(match[0], match, sayfa_numarasi) for match in matches])

    return kodlar


# Kullanım örneği
Tk().withdraw()
pdf_yolu = filedialog.askopenfilename(title="PDF Dosyasını Seç", filetypes=[("PDF Files", "*.pdf")])

if not pdf_yolu:
    print("PDF dosyası seçilmedi.")
else:
    sorular = pdf_den_kodlari_al(pdf_yolu)

    # Verileri DataFrame'e çevir
    data = {'İlk Harf': [kod[0] for kod in sorular],
            'Parantez İçi Veri': [kod[1] for kod in sorular],
            'Sayfa Numarası': [kod[2] for kod in sorular]}

    df = pd.DataFrame(data)

    # İlk harfe göre gruplama yaparak kaçıncı olduğunu belirten bir sütunu ekleyin
    df['İlk Harf Kaçıncı'] = df.groupby('İlk Harf').cumcount() + 1

    # Sütun sırasını düzenle
    df = df[['İlk Harf', 'İlk Harf Kaçıncı', 'Parantez İçi Veri', 'Sayfa Numarası']]

    # Excel dosyasını yazdırma
    excel_yolu = os.path.splitext(pdf_yolu)[0] + '_sorular.xlsx'
    df.to_excel(excel_yolu, index=False)
    print(f"Veriler '{excel_yolu}' dosyasına başarıyla yazıldı.")
