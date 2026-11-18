# SIPD TPP Migrator 📊

Automasi ekstraksi dan migrasi data **TPP (Tambahan Penghasilan Pegawai)** PNS dari laporan cetak Excel (tamsil) ke dalam template resmi **SIPD-RI** (`SIPD-Template-Data-TPP-PNSi.xlsx`).

Proyek ini menyediakan dua metode migrasi Python yang cepat, akurat, dan aman untuk memindahkan nominal TPP serta potongan pegawai tanpa perlu entri manual satu per satu.

---

## 🚀 Fitur & Metode Migrasi

### 1. Metode Pencocokan NIP (`migrate_match_nip.py`) — *Direkomendasikan*
- **Skenario Penggunaan:** Template SIPD sudah memiliki daftar pegawai (NIP & Nama), dan Anda hanya ingin mengisi nominal TPP serta potongannya.
- **Cara Kerja:**
  - Script membaca NIP 18 digit dari template SIPD baris demi baris.
  - Mencari NIP tersebut pada file sumber tamsil.
  - Mengisi **hanya** kolom target (TPP Beban Kerja, Kondisi Kerja, BPJS, PPh).
  - Kolom lain di template SIPD tetap aman dan tidak berubah.
  - Menampilkan laporan verifikasi: NIP mana yang cocok dan NIP mana yang tidak ditemukan di file sumber.

### 2. Metode Sekuensial (`migrate_sequential.py`)
- **Skenario Penggunaan:** Template SIPD masih berupa template kosong bawaan sistem (hanya berisi baris contoh), dan Anda ingin mengekstrak semua PNS dari file tamsil lalu mengisinya dari baris ke-2 secara berurutan.
- **Cara Kerja:**
  - Membersihkan baris sample bawaan template SIPD (baris 2–9).
  - Mengekstrak data PNS (NIP, Nama, Beban Kerja, Kondisi Kerja, BPJS 4%, BPJS 1%, PPh 21) dari file sumber.
  - Menuliskan data pegawai baru secara berurutan ke template SIPD.

---

## 📋 Pemetaan Kolom (Column Mapping)

| Data / Keterangan | Kolom Sumber (Tamsil Cetak) | Index Pandas (0-based) | Kolom Target SIPD | Index OpenPyXL (1-based) |
| :--- | :--- | :--- | :--- | :--- |
| **NIP Pegawai** | Kolom B (Teks gabungan) | Index 1 | Kolom A | Kolom 1 |
| **Nama Pegawai** | Kolom B (Baris pertama) | Index 1 | Kolom B | Kolom 2 |
| **TPP Beban Kerja** | Kolom L (*Beban Kerja Perolehan*) | Index 11 | Kolom P | Kolom 16 |
| **TPP Kondisi Kerja**| Kolom J (*Kondisi Kerja Perolehan*) | Index 9 | Kolom R | Kolom 18 |
| **Iuran JKN (BPJS 4%)**| Kolom O (*BPJS 4%*) | Index 14 | Kolom V | Kolom 22 |
| **Potongan IWP (BPJS 1%)**| Kolom Q (*BPJS 1%*) | Index 16 | Kolom AB | Kolom 28 |
| **Potongan PPh 21** | Kolom R (*PPh*) | Index 17 | Kolom AC | Kolom 29 |

---

## 🛠️ Instalasi & Prasyarat

Pastikan Anda telah menginstal **Python 3.8+**.

1. **Clone repository ini:**
   ```bash
   git clone https://github.com/aarosenn/sipd-tpp-migrator.git
   cd sipd-tpp-migrator
   ```

2. **Instal pustaka yang dibutuhkan:**
   ```bash
   pip install -r requirements.txt
   ```

---

## ⚙️ Petunjuk Penggunaan & Konfigurasi

1. Siapkan 2 file Excel berikut di folder yang sama:
   - File sumber tamsil cetak (misal: `sumber_tamsil_cetak_xxx.xlsx`)
   - File template resmi SIPD (misal: `SIPD-Template-Data-TPP-PNSi.xlsx`)
2. Buka script yang ingin digunakan (`migrate_match_nip.py` atau `migrate_sequential.py`).
3. Sesuaikan konfigurasi nama file pada bagian atas script:
   ```python
   # Ganti dengan nama file Anda:
   SOURCE   = "sumber_tamsil_cetak_xxx.xlsx"      # File sumber tamsil
   TEMPLATE = "SIPD-Template-Data-TPP-PNSi.xlsx"  # Template SIPD
   OUTPUT   = "SIPD-TPP-PNSi-Hasil-Terisi.xlsx"   # File hasil output
   ```
4. Jalankan script melalui terminal/command prompt:
   ```bash
   # Jika menggunakan Metode 1 (Pencocokan NIP):
   python migrate_match_nip.py

   # Atau jika menggunakan Metode 2 (Sekuensial):
   python migrate_sequential.py
   ```
5. File hasil output Excel akan otomatis dibuat dan siap diunggah ke SIPD-RI.

---

## 🔒 Keamanan & Perlindungan Data Sensitif

- Data kepegawaian (NIP, Nama) dan data keuangan (nominal penghasilan, pajak) merupakan **informasi rahasia/sensitif**.
- Repository ini secara bawaan telah menyertakan `.gitignore` untuk mengecualikan semua file `.xlsx`, `.xls`, dan `.csv`.
- **Jangan pernah meng-commit file Excel data riil ke dalam GitHub.** Selalu simpan file data di komputer lokal Anda.
