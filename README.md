# Generator Modul Ajar (AI - NVIDIA Build)

Aplikasi Streamlit untuk menyusun **Modul Ajar** (identitas, identifikasi
peserta didik, desain pembelajaran, kegiatan per pertemuan, asesmen, materi
ajar, tindak lanjut) memakai model AI dari [build.nvidia.com](https://build.nvidia.com),
lalu diunduh langsung sebagai file **.docx**.

## 1. Instalasi

```bash
pip install -r requirements.txt
```

## 2. Simpan API Key (jangan taruh di kode)

1. Buat akun / dapatkan API key di https://build.nvidia.com (klik model yang
   dipakai → "Get API Key").
2. Buat folder `.streamlit` di root proyek, lalu buat file `secrets.toml` di
   dalamnya (lihat contoh di `secrets.toml.example`):

```toml
NVIDIA_API_KEY = "nvapi-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

Struktur folder:

```
modul-ajar-streamlit/
├── app.py
├── requirements.txt
└── .streamlit/
    └── secrets.toml   <-- kamu buat sendiri, isi API key asli
```

## 3. Jalankan

```bash
streamlit run app.py
```

## Catatan tentang model

Kode ini memanggil model dengan nama persis:

```
nvidia/nemotron-3-ultra-550b-a55b
```

lewat endpoint OpenAI-compatible NVIDIA (`https://integrate.api.nvidia.com/v1`).
Jika nama model tersebut belum/tidak tersedia di akun NVIDIA Build kamu saat
mencoba, cek nama model yang valid di halaman katalog build.nvidia.com dan
ganti nilai `MODEL_NAME` di bagian atas `app.py`.

## Saat deploy ke Streamlit Community Cloud

Isi `NVIDIA_API_KEY` di menu **App settings → Secrets**, dengan format TOML
yang sama seperti `secrets.toml.example` — bukan ditulis di file `app.py`.
