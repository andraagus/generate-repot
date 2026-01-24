import pandas as pd
import glob
import os

# 1. Setup Path
script_dir = os.path.dirname(os.path.abspath(__file__))
print(f"[*] Menjalankan script di: {script_dir}")

# --- input user
periode_input = input("Masukan Tahun dan Bulan (Contoh: 2024-05): ").strip()

# PERBAIKAN: Hapus .astype().str.strip() karena ini variabel string biasa
output_filename = os.path.join(script_dir, f'Hasil_Laporan_penambahan_loan_baru_{periode_input}.xlsx')

# 2. Cari file KRP
search_pattern = os.path.join(script_dir, 'KRP*.xls*')
krp_files = glob.glob(search_pattern)

if not krp_files:
    print(f"[!] Error: File KRP tidak ditemukan di {script_dir}")
    exit()

file_krp = krp_files[0]
print(f"[*] Memproses file: {os.path.basename(file_krp)}")

# Membaca sebagai string untuk keamanan data awal
df_krp = pd.read_excel(file_krp, dtype=str)

# 3. filter data berdasarkan periode_input
if 'tanggalAwal' in df_krp.columns:
    # Konversi ke datetime untuk filter yang akurat
    temp_date = pd.to_datetime(df_krp['tanggalAwal'], errors='coerce')
    df_filtered = df_krp[temp_date.dt.strftime('%Y-%m') == periode_input].copy()
else:
    print("[!] Kolom 'tanggalAwal' tidak ditemukan di file KRP!")
    exit()

if df_filtered.empty:
    print(f"[!] Tidak ada data loan baru untuk periode {periode_input}.")
    exit()

# 4. Seleksi dan Rename kolom
# PERBAIKAN: Gunakan format dictionary { 'asal': 'tujuan' }
mapping_kolom = {
    'NAMA_CUSTOMER':'NAMA NASABAH',
    'nomorRekening':'NOMOR REKENING',
    'PRODUK':'PRODUK',
    'plafon':'PLAFOND',
    'tanggalMulai':'TANGGAL MULAI',
    'sukuBungaPersentaseImbalanBulanLaporan':'BUNGA PINJAMAN'
}

# Ambil kolom yang benar-benar ada saja
kolom_tersedia = [col for col in mapping_kolom.keys() if col in df_filtered.columns]
df_new = df_filtered[kolom_tersedia].copy()
df_new = df_new.rename(columns=mapping_kolom)

# 5. Tambahkan kolom "No" di posisi paling depan
df_new.insert(0, 'No', range(1, len(df_new) + 1))

# 6. Pembersihan Akhir (Semua dipaksa jadi String & Strip)
df_new = df_new.astype(str)
for col in df_new.columns:
    df_new[col] = df_new[col].str.strip().replace('nan', '')

# 7. Simpan ke Excel
df_new.to_excel(output_filename, index=False)

print(f"\n[V] Berhasil! Ditemukan {len(df_new)} data.")
print(f"[V] File disimpan: {os.path.basename(output_filename)}")