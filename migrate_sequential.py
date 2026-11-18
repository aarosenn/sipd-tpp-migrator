"""
Migrasi Data TPP PNS ke Template SIPD-RI (Metode Sekuensial)

Script ini mengekstrak data PNS dari file sumber cetak tamsil,
mengosongkan baris contoh bawaan template SIPD, lalu menuliskan
data NIP, Nama, TPP, dan Potongan secara berurutan mulai dari baris ke-2.
"""

import re
import openpyxl
import pandas as pd


# ============================================================
# KONFIGURASI FILE
# ============================================================
# Petunjuk: Ganti string di bawah ini sesuai dengan nama file Excel Anda.
# Pastikan file diletakkan di direktori yang sama dengan script ini.
SOURCE   = "sumber_tamsil_cetak_xxx.xlsx"      # Ganti dengan file sumber data cetak tamsil
TEMPLATE = "SIPD-Template-Data-TPP-PNSi.xlsx"  # Ganti dengan file template SIPD asli
OUTPUT   = "SIPD-TPP-PNS-Terisi.xlsx"          # Nama file hasil keluaran


def migrate_tpp_pns(source_file, template_file, output_file):
    print("=" * 65)
    print("  MIGRASI DATA TPP PNS -> SIPD (METODE SEKUENSIAL)")
    print("=" * 65)
    print(f"  Sumber   : {source_file}")
    print(f"  Template : {template_file}")
    print(f"  Output   : {output_file}")
    print("=" * 65)

    # 1. Baca data sumber
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

    # 2. Filter dan ekstrak data PNS (baris data dimulai dari baris index 6)
    print("\n[2/4] Mengekstrak data PNS...")
    pns_rows = []
    for idx in range(6, len(df_src)):
        cell_info = df_src.iloc[idx, 1]
        if pd.isna(cell_info):
            continue

        text = str(cell_info).strip()
        # Filter khusus pegawai berstatus PNS
        if "PNS" in text:
            lines = [line.strip() for line in text.split("\n") if line.strip()]
            nama = lines[0]

            # Ambil NIP 18 digit
            nip_match = re.search(r"\d{18}", text)
            nip = (
                nip_match.group(0)
                if nip_match
                else (lines[1] if len(lines) > 1 else "")
            )

            # Pemetaan kolom sumber:
            kondisi_kerja = df_src.iloc[idx, 9]   # Kolom J: Kondisi Kerja Perolehan
            beban_kerja   = df_src.iloc[idx, 11]  # Kolom L: Beban Kerja Perolehan
            bpjs_4        = df_src.iloc[idx, 14]  # Kolom O: BPJS 4%
            bpjs_1        = df_src.iloc[idx, 16]  # Kolom Q: Potongan BPJS 1%
            pph           = df_src.iloc[idx, 17]  # Kolom R: Potongan PPh

            pns_rows.append(
                {
                    "nip": nip,
                    "nama": nama,
                    "beban_kerja": beban_kerja,
                    "kondisi_kerja": kondisi_kerja,
                    "bpjs_4": bpjs_4,
                    "bpjs_1": bpjs_1,
                    "pph": pph,
                }
            )

    print(f"      Ditemukan {len(pns_rows)} pegawai PNS.")
    if not pns_rows:
        print("  [PERINGATAN] Tidak ada data PNS yang ditemukan.")
        return

    # 3. Buka file template SIPD menggunakan openpyxl
    print("\n[3/4] Membuka template SIPD dan membersihkan data sample...")
    try:
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
    except FileNotFoundError:
        print(f"  [ERROR] File template tidak ditemukan: '{template_file}'")
        return
    except Exception as e:
        print(f"  [ERROR] Gagal membuka template: {e}")
        return

    # Kosongkan baris contoh bawaan template (baris 2 sampai sebelum catatan keterangan)
    # Catatan SIPD umumnya mulai di baris 10
    start_row = 2
    for r in range(start_row, 10):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).value = None

    # 4. Tulis data ke kolom target SIPD:
    # Col 1  (A)  : NIP Pegawai
    # Col 2  (B)  : Nama Pegawai
    # Col 16 (P)  : TPP Beban Kerja
    # Col 18 (R)  : TPP Kondisi kerja
    # Col 22 (V)  : Iuran Jaminan Kesehatan (BPJS 4%)
    # Col 28 (AB) : Potongan IWP (BPJS 1%)
    # Col 29 (AC) : Potongan PPh 21
    print("\n[4/4] Menulis data pegawai ke template SIPD...")
    current_row = start_row
    for item in pns_rows:
        ws.cell(row=current_row, column=1, value=str(item["nip"]))
        ws.cell(row=current_row, column=2, value=item["nama"])
        ws.cell(row=current_row, column=16, value=item["beban_kerja"])
        ws.cell(row=current_row, column=18, value=item["kondisi_kerja"])
        ws.cell(row=current_row, column=22, value=item["bpjs_4"])
        ws.cell(row=current_row, column=28, value=item["bpjs_1"])
        ws.cell(row=current_row, column=29, value=item["pph"])
        current_row += 1

    try:
        wb.save(output_file)
    except PermissionError:
        print(f"\n  [ERROR] File '{output_file}' sedang terbuka di Excel. Tutup file dan ulangi.")
        return
    except Exception as e:
        print(f"\n  [ERROR] Gagal menyimpan file: {e}")
        return

    print("=" * 65)
    print(f"  SELESAI! {len(pns_rows)} data PNS berhasil dipindahkan ke '{output_file}'.")
    print("=" * 65)


# Eksekusi script
if __name__ == "__main__":
    migrate_tpp_pns(SOURCE, TEMPLATE, OUTPUT)
