import pandas as pd
import glob
import os

# 1. Setup Path
script_dir = os.path.dirname(os.path.abspath(__file__))
print(f"[*] Menjalankan script di: {script_dir}")

# --- Input User ---
periode_input = input("Masukkan Batas Periode Tanggal Mulai (Contoh 2024-05): ").strip()

output_filename = os.path.join(script_dir, f'Hasil_Laporan_Lunas_F01_{periode_input}.xlsx')

# 2. Cari File Sumber
file_f01_list = glob.glob(os.path.join(script_dir, '*F01*.xls*'))
path_reff = os.path.join(script_dir, 'reff-sektor-ekonomi.xlsx')

if not file_f01_list or not os.path.exists(path_reff):
    print("[!] Error: File F01 atau Reff-Sektor-Ekonomi tidak ditemukan!")
    exit()

print(f"[*] Membaca file F01: {os.path.basename(file_f01_list[0])}")

# 3. Baca Data (Semua sebagai String)
df_f01 = pd.read_excel(file_f01_list[0], dtype=str)
df_reff = pd.read_excel(path_reff, dtype=str)

# 4. Filter Tanggal & Status Lunas
# Konversi Tanggal Mulai untuk perbandingan (YYYY-MM-DD)
df_f01['dt_mulai_temp'] = pd.to_datetime(df_f01['Tanggal Mulai'], errors='coerce')
limit_date = pd.to_datetime(periode_input, format='%Y-%m')

# Filter: 
# 1. Tanggal Mulai < periode_input
# 2. Kode Kondisi == '02' (Gunakan zfill untuk memastikan format dua digit)
df_lunas = df_f01[
    (df_f01['dt_mulai_temp'] < limit_date) & 
    (df_f01['Kode Kondisi'].str.strip().str.zfill(2) == '02')
].copy()

if df_lunas.empty:
    print(f"[!] Tidak ada data dengan Kode Kondisi '02' sebelum {periode_input}")
    exit()

# 5. Merge dengan Referensi Sektor Ekonomi
df_reff['Sandi Referensi'] = df_reff['Sandi Referensi'].str.strip()
df_lunas['Kode Sektor Ekonomi'] = df_lunas['Kode Sektor Ekonomi'].str.strip()

df_merged = pd.merge(
    df_lunas,
    df_reff[['Sandi Referensi', 'Definisi']],
    left_on='Kode Sektor Ekonomi',
    right_on='Sandi Referensi',
    how='left'
)

# 6. Susun Kolom Sesuai Permintaan
# Format: 'Nama Kolom di F01': 'Nama Kolom Baru di Laporan'
mapping_kolom = {
    'No Rekening Fasilitas': 'Norekeing',
    'Tanggal Awal Kredit': 'tanggalawal',
    'Tanggal Kondisi': 'tanggal kondisi',
    'No CIF Debitur': 'cif',
    'Keterangan': 'customer name', 
    'Plafon Awal': 'plafond',
    'Baki Debet': 'pokok terakhir',
    'Kode Jenis Penggunaan': 'jenis penggunaan',
    'Kode Sektor Ekonomi': 'sektor ekonomi',
    'Definisi': 'definis'
}

# Pastikan hanya mengambil kolom yang benar-benar ada
kolom_final_keys = [col for col in mapping_kolom.keys() if col in df_merged.columns]
df_final = df_merged[kolom_final_keys].copy()
df_final = df_final.rename(columns=mapping_kolom)

# 7. Tambahkan Kolom No & Pembersihan String Final
df_final.insert(0, 'No', range(1, len(df_final) + 1))
df_final = df_final.astype(str)

for col in df_final.columns:
    # Menghapus spasi dan mengganti teks 'nan' (hasil dari merge/empty cell) menjadi string kosong
    df_final[col] = df_final[col].str.strip().replace(['nan', 'None', 'NaT'], '')

# 8. Simpan ke Excel Baru
df_final.to_excel(output_filename, index=False)

print(f"---")
print(f"[V] Berhasil memproses {len(df_final)} data lunas.")
print(f"[V] File disimpan di: {output_filename}")