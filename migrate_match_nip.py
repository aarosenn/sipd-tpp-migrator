"""
Migrasi Data TPP PNS ke Template SIPD-RI (Metode Pencocokan NIP)

Script ini mencocokkan NIP pada template SIPD dengan data sumber cetak tamsil.
Hanya kolom nominal TPP dan potongan yang diisi, kolom lain pada template
tidak disentuh sehingga aman.
"""

import re
import openpyxl
import pandas as pd


# ============================================================
# KONFIGURASI FILE
# ============================================================
# Petunjuk: Ganti string di bawah ini sesuai dengan nama file Excel Anda.
# Pastikan file diletakkan di direktori yang sama dengan script ini.
SOURCE   = "sumber_tamsil_cetak_xxx.xlsx"         # Ganti dengan file sumber data cetak tamsil
TEMPLATE = "SIPD-Template-Data-TPP-PNSi.xlsx"     # Ganti dengan file template SIPD asli
OUTPUT   = "SIPD-TPP-PNSi-Hasil-Terisi.xlsx"      # Nama file hasil keluaran


# ============================================================
# PEMETAAN KOLOM SUMBER -> KOLOM TARGET SIPD
#
# File Sumber (tamsil cetak):
#   Kolom J (index 9)  = Kondisi Kerja Perolehan
#   Kolom L (index 11) = Beban Kerja Perolehan
#   Kolom O (index 14) = BPJS 4%
#   Kolom Q (index 16) = BPJS 1%
#   Kolom R (index 17) = PPh
#
# File Template SIPD (1-based openpyxl):
#   Kolom P (16)  = TPP Beban Kerja
#   Kolom R (18)  = TPP Kondisi Kerja
#   Kolom V (22)  = Iuran Jaminan Kesehatan (BPJS 4%)
#   Kolom AB (28) = Potongan IWP (BPJS 1%)
#   Kolom AC (29) = Potongan PPh 21
# ============================================================

# Kolom sumber (0-based pandas index)
SRC_COL = {
    "kondisi_kerja": 9,   # Kolom J
    "beban_kerja":   11,  # Kolom L
    "bpjs_4":        14,  # Kolom O
    "bpjs_1":        16,  # Kolom Q
    "pph":           17,  # Kolom R
}

# Kolom target di SIPD template (1-based openpyxl index)
# Key = nomor kolom SIPD, Value = nama field dari sumber
TARGET_COL = {
    16: "beban_kerja",    # Kolom P  - TPP Beban Kerja
    18: "kondisi_kerja",  # Kolom R  - TPP Kondisi Kerja
    22: "bpjs_4",         # Kolom V  - Iuran Jaminan Kesehatan (BPJS 4%)
    28: "bpjs_1",         # Kolom AB - Potongan IWP (BPJS 1%)
    29: "pph",            # Kolom AC - Potongan PPh 21
}

# Kolom NIP & Nama di template SIPD (untuk pencocokan)
COL_NIP_SIPD = 1   # Kolom A
COL_NAMA_SIPD = 2  # Kolom B


# ============================================================
# FUNGSI BANTU
# ============================================================
def safe_val(df, row_idx, col_idx):
    """Ambil nilai dari DataFrame, kembalikan None jika kosong/NaN."""
    val = df.iloc[row_idx, col_idx]
    return None if pd.isna(val) else val


def ekstrak_nip(text):
    """Ambil NIP 18 digit dari teks sel kolom B file sumber."""
    match = re.search(r"\d{18}", text)
    return match.group(0) if match else None


# ============================================================
# FUNGSI UTAMA
# ============================================================
def migrate_tpp_pnsi(source_file, template_file, output_file):
    print("=" * 65)
    print("  MIGRASI DATA TPP PNS -> SIPD (METODE PENCOCOKAN NIP)")
    print("=" * 65)
    print(f"  Sumber   : {source_file}")
    print(f"  Template : {template_file}")
    print(f"  Output   : {output_file}")
    print("=" * 65)

    # ----------------------------------------------------------
    # LANGKAH 1: Baca file sumber
    # ----------------------------------------------------------
    print("\n[1/4] Membaca file sumber...")
    try:
        df_src = pd.read_excel(source_file, header=None)
        print(f"      Berhasil. Total baris: {len(df_src)}")
    except FileNotFoundError:
        print(f"  [ERROR] File tidak ditemukan: '{source_file}'")
        return
    except Exception as e:
        print(f"  [ERROR] Gagal membaca file sumber: {e}")
        return

    # ----------------------------------------------------------
    # LANGKAH 2: Ekstrak semua data PNS dari file sumber
    #            Baris data dimulai dari index 6 (baris ke-7 Excel)
    # ----------------------------------------------------------
    print("\n[2/4] Mengekstrak data PNS dari file sumber...")
    data_sumber = {}  # dict: NIP -> dict nilai kolom

    for idx in range(6, len(df_src)):
        cell_b = df_src.iloc[idx, 1]
        if pd.isna(cell_b):
            continue

        text = str(cell_b).strip()
        if "PNS" not in text:
            continue

        nip = ekstrak_nip(text)
        if not nip:
            print(f"      [SKIP] Baris {idx+1}: NIP tidak ditemukan -> {text[:40]}")
            continue

        data_sumber[nip] = {
            field: safe_val(df_src, idx, col_idx)
            for field, col_idx in SRC_COL.items()
        }

    print(f"      Ditemukan {len(data_sumber)} pegawai PNS di file sumber.")

    if not data_sumber:
        print("  [PERINGATAN] Tidak ada data PNS ditemukan. Script dihentikan.")
        return

    # ----------------------------------------------------------
    # LANGKAH 3: Buka template SIPD
    # ----------------------------------------------------------
    print("\n[3/4] Membuka template SIPD...")
    try:
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        print(f"      Sheet: '{ws.title}' | {ws.max_row} baris x {ws.max_column} kolom")
    except FileNotFoundError:
        print(f"  [ERROR] File template tidak ditemukan: '{template_file}'")
        return
    except Exception as e:
        print(f"  [ERROR] Gagal membuka template: {e}")
        return

    # ----------------------------------------------------------
    # LANGKAH 4: Cocokkan NIP dan isi kolom target
    #            Kolom lain yang ada di template TIDAK disentuh.
    # ----------------------------------------------------------
    print("\n[4/4] Mencocokkan NIP dan mengisi data ke template...")
    print(f"      Kolom yang akan diisi: {sorted(TARGET_COL.keys())} (1-based)")
    print()

    cocok = 0
    tidak_cocok = []

    for row in range(2, ws.max_row + 1):
        nip_sipd = ws.cell(row=row, column=COL_NIP_SIPD).value
        nama_sipd = ws.cell(row=row, column=COL_NAMA_SIPD).value

        if not nip_sipd:
            continue

        nip_sipd_str = str(nip_sipd).strip()

        if nip_sipd_str in data_sumber:
            data = data_sumber[nip_sipd_str]

            # Isi HANYA kolom yang ada di TARGET_COL
            for col_num, field_name in TARGET_COL.items():
                ws.cell(row=row, column=col_num, value=data[field_name])

            cocok += 1
            print(f"  [OK] Baris {row:2d} | NIP: {nip_sipd_str} | {nama_sipd}")
        else:
            tidak_cocok.append((row, nip_sipd_str, nama_sipd))
            print(f"  [--] Baris {row:2d} | NIP: {nip_sipd_str} | {nama_sipd} <- tidak ada di sumber")

    # ----------------------------------------------------------
    # Simpan output
    # ----------------------------------------------------------
    try:
        wb.save(output_file)
    except PermissionError:
        print(f"\n  [ERROR] File '{output_file}' sedang terbuka di Excel.")
        print("         Tutup file tersebut lalu jalankan script lagi.")
        return
    except Exception as e:
        print(f"\n  [ERROR] Gagal menyimpan file: {e}")
        return

    # ----------------------------------------------------------
    # Ringkasan
    # ----------------------------------------------------------
    print()
    print("=" * 65)
    print("  SELESAI!")
    print("=" * 65)
    print(f"  Total NIP di template  : {cocok + len(tidak_cocok)}")
    print(f"  Berhasil dicocokkan    : {cocok}")
    print(f"  Tidak ada di sumber    : {len(tidak_cocok)}")
    if tidak_cocok:
        print(f"\n  NIP yang tidak cocok:")
        for baris, nip, nama in tidak_cocok:
            print(f"    Baris {baris}: {nip} | {nama}")
    print()
    print(f"  Kolom yang diisi       : P(16), R(18), V(22), AB(28), AC(29)")
    print(f"  Kolom lain di template : TIDAK diubah (aman)")
    print(f"  File output            : {output_file}")
    print("=" * 65)


# ============================================================
# EKSEKUSI
# ============================================================
if __name__ == "__main__":
    migrate_tpp_pnsi(SOURCE, TEMPLATE, OUTPUT)
