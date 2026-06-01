# 🚀 Panduan Deploy Excel→PDF API ke Render.com (GRATIS)

## Isi Folder `excel-pdf-api`:
```
excel-pdf-api/
├── app.py           ← Script Python utama
├── requirements.txt ← Library yang dibutuhkan
├── render.yaml      ← Konfigurasi Render.com
└── Procfile         ← Perintah start server
```

---

## Step 1 — Upload ke GitHub (wajib)

1. Buka **github.com** → Login atau daftar gratis
2. Klik **"New repository"** → nama: `excel-pdf-api`
3. Klik **"uploading an existing file"**
4. Upload semua file dari folder `excel-pdf-api`
5. Klik **"Commit changes"** → selesai!

---

## Step 2 — Deploy ke Render.com (gratis)

1. Buka **render.com** → daftar pakai akun GitHub
2. Klik **"New +"** → **"Web Service"**
3. Pilih repository **excel-pdf-api** dari GitHub kamu
4. Isi pengaturan:
   - **Name**: excel-pdf-api
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Plan**: Free ✅
5. Klik **"Create Web Service"**
6. Tunggu ~3 menit sampai deploy selesai
7. Kamu dapat URL seperti: `https://excel-pdf-api.onrender.com`

---

## Step 3 — Update URL di Website

Buka file `index.html` di website FileToolPro kamu.
Cari baris ini:
```javascript
var API_URL = 'https://excel-pdf-api.onrender.com/convert/excel-to-pdf';
```

Ganti dengan URL asli dari Render.com kamu:
```javascript
var API_URL = 'https://NAMA-PROJECT-KAMU.onrender.com/convert/excel-to-pdf';
```

Lalu upload ulang `index.html` ke Netlify.

---

## ⚠️ Catatan Render.com Free Plan:

- Server **"tidur"** setelah 15 menit tidak ada request
- Request pertama bisa lambat ~30 detik (server bangun)
- Request berikutnya normal cepat
- Kalau mau selalu cepat → upgrade ke Render Starter $7/bulan

---

## ✅ Test API

Buka browser, ketik:
```
https://excel-pdf-api.onrender.com/health
```

Kalau muncul `{"status": "ok"}` → API berjalan dengan baik!

---

## Keuntungan Python API vs Browser-only:

| Fitur | Browser JS | Python API |
|-------|-----------|------------|
| Warna sel Excel | ⚠️ Sebagian | ✅ Semua |
| Font bold/italic | ⚠️ Sebagian | ✅ Semua |
| Lebar kolom | ✅ | ✅ |
| Tinggi baris | ✅ | ✅ |
| Theme colors | ❌ | ✅ |
| Format angka (Rp) | ⚠️ | ✅ |
| Kecepatan | Fast | Medium |
| Butuh internet | Tidak | Ya |

Python API menggunakan **openpyxl** (library resmi untuk baca Excel)
dan **ReportLab** (library profesional untuk generate PDF).
