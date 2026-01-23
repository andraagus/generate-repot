import pandas as pd
import glob

# Baca referensi sektor ekonomi
reff = pd.read_excel('Reff-sektor-ekonomi.xlxs')
reff = reff.rename(columns=lambda x: x.strip())
reff_dict = dict(zip(reff['Sandi Referensi'], reff['Definisi']))

# Temukan file KRP (bisa lebih dari satu, ambil semua yang sesuai pola)
krp_files = glob.glob('KRP - *.xls')

# Kolom yang diambil dari file KRP
kolom_ambil = [
    'nomorRekening', 'NAMA_CUSTOMER', 'jenisKreditPembiayaan', 'tanggalAwal',
    'tanggalMulai', 'tanggalJatuhTempo', 'sukuBungaPersentaseImbalanBulanLaporan',
    'jenisPenggunaan', 'sektorEkonomi', 'kualitas', 'plafon', 'bakiDebet', 'jumlah', 'NIK', 'PARTNER'
]

# Gabungkan semua file KRP jika lebih dari satu
df_list = []
for file in krp_files:
    df = pd.read_excel(file, dtype=str)
    df = df.rename(columns=lambda x: x.strip())
    df_list.append(df)
krp = pd.concat(df_list, ignore_index=True)

# Ambil kolom yang dibutuhkan
krp = krp[kolom_ambil]

# Tambahkan kolom Definisi dari referensi sektor ekonomi
krp['Definisi'] = krp['sektorEkonomi'].map(reff_dict)

# Urutkan kolom sesuai permintaan
kolom_final = [
    'nomorRekening', 'NAMA_CUSTOMER', 'jenisKreditPembiayaan', 'tanggalAwal',
    'tanggalMulai', 'tanggalJatuhTempo', 'sukuBungaPersentaseImbalanBulanLaporan',
    'jenisPenggunaan', 'sektorEkonomi', 'Definisi', 'kualitas', 'plafon', 'bakiDebet', 'jumlah', 'NIK'
]
krp = krp[kolom_final + ['PARTNER']]

# Kelompokkan per PARTNER dan simpan ke file Excel, satu sheet per PARTNER
with pd.ExcelWriter('output_per_partner.xlsx', engine='openpyxl') as writer:
    for partner, group in krp.groupby('PARTNER'):
        group = group.drop(columns=['PARTNER'])
        group.to_excel(writer, sheet_name=str(partner)[:31], index=False)

print("File output_per_partner.xlsx berhasil dibuat.")