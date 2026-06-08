# 📄 Teslim Tesellüm Oluşturucu — Web Versiyonu

Excel listesinden otomatik teslim-tesellüm belgesi (PDF) üretir.  
**LibreOffice gerekmez** — tamamen Python ile çalışır.

---

## 🚀 Streamlit Cloud'a Deploy (Ücretsiz, 5 Dakika)

### 1. GitHub'a yükle

```bash
git init
git add .
git commit -m "ilk yükleme"
git remote add origin https://github.com/KULLANICI_ADIN/teslim-web.git
git push -u origin main
```

### 2. Streamlit Cloud'da aç

1. [share.streamlit.io](https://share.streamlit.io) adresine git
2. **"New app"** butonuna tıkla
3. GitHub reposunu seç
4. **Main file path:** `app.py`
5. **Deploy!**

Birkaç dakika sonra uygulamanız şu formatta bir URL'de yayında olur:  
`https://KULLANICI.streamlit.app`

Bu linki muhasebe ekibinizle paylaşın — herhangi bir kurulum gerektirmez.

---

## 💻 Yerel Çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 📁 Proje Yapısı

```
teslim_web/
├── app.py               # Ana uygulama
├── requirements.txt     # Bağımlılıklar
├── templates/
│   ├── akakce_alis.xlsx
│   ├── akakce_satis.xlsx
│   └── ...              # Diğer firma şablonları buraya
└── README.md
```

---

## ➕ Yeni Firma Şablonu Eklemek

`templates/` klasörüne `{firma}_{alis|satis}.xlsx` formatında ekleyin.  
Örnek: `cevher_alis.xlsx`, `altinsa_satis.xlsx`

`app.py` dosyasındaki `FIRMA_LISTESI` listesine firma adını da ekleyin:

```python
FIRMA_LISTESI = ["akakçe", "cevher", "çıtır", "altınsa"]
```

---

## 📦 Çıktı (ZIP)

```
teslim_tesellum_ciktilari.zip
├── pdfler/
│   ├── FATURA001.pdf
│   ├── FATURA002.pdf
│   └── ...
└── tum_pdfler.pdf        # (veya gruplu_pdfler/ klasörü)
```
