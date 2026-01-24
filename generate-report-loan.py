import pandas as pd
import glob
import os

# 1. Setup Path
script_dir = os.path.dirname(os.path.abspath(__file__))
print(f"[*] Menjalankan script di: {script_dir}")

# Path untuk file referensi
path_reff = os.path.join(script_dir, 'reff-sektor-ekonomi.xlsx')
output_filename = os.path.join(script_dir, 'Hasil_Laporan__loan_Per_Partner.xlsx')

# 2. Load file Referensi
if not os.path.exists(path_reff):
    print(f"[!] Error: File referensi tidak ditemukan di: {path_reff}")
    exit()

df_reff = pd.read_excel(path_reff, dtype=str)
# Pastikan kunci join adalah string agar sinkron
df_reff['Sandi Referensi'] = df_reff['Sandi Referensi'].astype(str).str.strip()

# 3. Cari dan Load file KRP
# Mencari .xls atau .xlsx yang diawali dengan KRP
search_pattern = os.path.join(script_dir, 'KRP*.xls*')
krp_files = glob.glob(search_pattern)

if not krp_files:
    print(f"[!] Error: File KRP tidak ditemukan di {script_dir}")
    print(f"Isi folder saat ini: {os.listdir(script_dir)}")
else:
    file_krp = krp_files[0]
    print(f"[*] Memproses file: {os.path.basename(file_krp)}")
    df_krp = pd.read_excel(file_krp, dtype=str)

    # Pastikan kunci join di KRP juga string
    if 'sektorEkonomi' in df_krp.columns:
        df_krp['sektorEkonomi'] = df_krp['sektorEkonomi'].astype(str).str.strip()
    else:
        print("[!] Kolom 'sektorEkonomi' tidak ditemukan di file KRP!")
        exit()

    # 4. Proses Merge (VLOOKUP)
    df_merged = pd.merge(
        df_krp, 
        df_reff[['Sandi Referensi', 'Definisi']], 
        left_on='sektorEkonomi', 
        right_on='Sandi Referensi', 
        how='left'
    )

    # 5. Seleksi kolom yang dibutuhkan
    kolom_pilihan = [
        'nomorRekening', 'NAMA_CUSTOMER', 'jenisKreditPembiayaan', 
        'tanggalAwal', 'tanggalMulai', 'tanggalJatuhTempo', 
        'sukuBungaPersentaseImbalanBulanLaporan', 'jenisPenggunaan', 
        'sektorEkonomi', 'Definisi', 'kualitas', 'plafon', 
        'bakiDebet', 'jumlah', 'NIK', 'jumlahHariTunggakan','PARTNER'
    ]
    
    # Filter hanya kolom yang benar-benar ada di hasil merge untuk menghindari error
    kolom_tersedia = [c for c in kolom_pilihan if c in df_merged.columns]
    df_final = df_merged[kolom_tersedia].copy()

    # 6. Export ke Excel per Sheet
    print(f"[*] Mengekspor data ke {output_filename}...")
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        if 'PARTNER' in df_final.columns:
            # Ambil partner unik, ganti NaN dengan 'Tanpa Partner'
            df_final['PARTNER'] = df_final['PARTNER'].fillna('Tanpa Partner')
            partners = df_final['PARTNER'].unique()
            
            for partner in partners:
                df_partner = df_final[df_final['PARTNER'] == partner]
                
                # Bersihkan nama sheet (max 31 karakter, hapus karakter terlarang)
                sheet_name = str(partner)[:31].replace('/', '_').replace('\\', '_')
                for char in ['*', ':', '/', '\\', '?', '[', ']']:
                    sheet_name = sheet_name.replace(char, '_')
                
                df_partner.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # Jika kolom PARTNER tidak ditemukan, simpan semua di satu sheet
            df_final.to_excel(writer, sheet_name='Data_Laporan', index=False)

    print("Status: Berhasil!")