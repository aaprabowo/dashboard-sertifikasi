# Dashboard Monitoring Terpadu — Panduan Setup

## Struktur File

```
dashboard-sertifikasi/
├── index.html          ← Dashboard utama (buka di browser)
├── watcher.py          ← Script auto-update dari Excel SharePoint
├── README.md           ← Panduan ini
└── data/
    └── dashboard_data.json   ← Dibuat otomatis oleh watcher.py
```

---

## Format Excel (Satu File, Banyak Sheet)

Satu file Excel dengan sheet terpisah per tim. Nama sheet harus sesuai:

| Nama Sheet    | Tim          |
|---------------|--------------|
| `JFKN`        | Tim JFKN     |
| `SAK`         | Tim SAK      |
| `AC`          | Tim AC       |
| `Beasiswa`    | Tim Beasiswa |
| `USKP`        | Tim USKP     |
| `PBJ`         | Tim PBJ      |
| `Tes Lainnya` | Tim Tes Lainnya |

Jika nama sheet berbeda, sesuaikan `SHEET_MAP` di `watcher.py`.

---

## Cara A: OneDrive Sync + Watcher Otomatis (Rekomendasi)

### Langkah 1 — Install dependensi (sekali saja)
```bash
pip install openpyxl watchdog
```

### Langkah 2 — Edit path file Excel di watcher.py
```python
# Ubah baris ini ke path lengkap file Excel Anda:
EXCEL_FILE = r"C:\Users\NamaAnda\OneDrive - NamaOrganisasi\Data Sertifikasi\Data_Monitoring.xlsx"
```

### Langkah 3 — Sesuaikan nama sheet (jika perlu)
```python
SHEET_MAP = {
    "JFKN":        "JFKN",       # ← nama tab sheet di Excel
    "SAK":         "SAK",
    "Tes Lainnya": "Tes Lainnya", # ← sesuaikan jika beda
    ...
}
```

### Langkah 4 — Jalankan watcher
```bash
python watcher.py
```

Output saat berhasil:
```
09:00:01 [INFO] File Excel : C:\...\Data_Monitoring.xlsx
09:00:01 [INFO] Sheet tersedia: ['JFKN', 'SAK', 'AC', 'Beasiswa', 'USKP', 'PBJ', 'Tes Lainnya']
09:00:02 [INFO] [JFKN] Diproses: 680 peserta, 578 lulus (85.0%)
...
09:00:03 [INFO] ✓ JSON diperbarui → data/dashboard_data.json
09:00:03 [INFO] Watcher aktif — memantau: C:\...\Data Sertifikasi
```

### Langkah 5 — Hosting di LAN
```bash
python -m http.server 3000
```
Akses dari komputer lain: `http://[IP-PC-SERVER]:3000`

---

## Cara B: Upload Manual (tanpa watcher)

Buka `index.html` → tab **Modul Operator** → drag & drop file Excel.
Semua sheet dideteksi dan diproses sekaligus secara otomatis.

---

## Kolom Wajib per Sheet

### JFKN / SAK / AC
| Kolom      | Status   |
|------------|----------|
| Status     | **Wajib** — nilai: `Lulus` / `Tidak Lulus` |
| Batch      | **Wajib** — untuk grafik tren |
| Unit Kerja | **Wajib** — untuk breakdown unit |
| Nama, NIP  | Wajib (tidak dikirim ke dashboard) |

### Beasiswa
Sama seperti di atas, plus kolom `Angkatan` (dipakai sebagai pengganti Batch).

### USKP / PBJ / Tes Lainnya
Sama seperti di atas, plus kolom `Periode` (dipakai sebagai pengganti Batch).

Nilai kolom **Status** yang dikenali sebagai lulus:
`Lulus`, `L`, `Passed`, `P`, `Ya`, `Y` (tidak case-sensitive)

---

## Keamanan Data

- **Data individu (Nama, NIP, dll) tidak pernah keluar dari Excel** — watcher hanya membaca dan menghitung agregat
- **Dashboard hanya menampilkan persentase dan total** — tidak ada data per-orang
- **Tidak ada koneksi internet** — semua proses di LAN kantor
- **Satu file Excel** — lebih mudah dikontrol aksesnya via permission SharePoint/OneDrive

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| "File Excel tidak ditemukan" | Cek path `EXCEL_FILE` di watcher.py, pastikan OneDrive sudah sync |
| Sheet tidak terbaca | Cek nama sheet di Excel harus persis sama dengan `SHEET_MAP` |
| "Kolom 'Status' tidak ditemukan" | Cek nama kolom di sheet tersebut |
| Dashboard tidak update | Pastikan watcher.py berjalan dan `data/dashboard_data.json` ada |
| Tidak bisa akses dari PC lain | Izinkan port 3000 di firewall Windows |
