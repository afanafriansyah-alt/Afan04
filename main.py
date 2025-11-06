# ================================================
# Program: Rekap Nilai Mahasiswa dari Excel
# Library: pandas dan openpyxl
# ================================================

import pandas as pd

# 1️. Import file Excel
file_path = "data_mahasiswa.xlsx"   # Pastikan nama file sesuai
data = pd.read_excel(file_path)

# 2️. Hitung rata-rata dengan bobot:
#    Tugas = 30%, UTS = 30%, UAS = 40%
data['Rata-rata'] = (data['Nilai 1'] * 0.3) + (data['Nilai 2'] * 0.3) + (data['Nilai 3'] * 0.4)

# 3️. Tambahkan kolom Status:
#    Jika rata-rata ≥ 75 → Lulus
#    Jika rata-rata < 75 → Tidak Lulus
data['Status'] = data['Rata-rata'].apply(lambda x: 'Lulus' if x >= 75 else 'Tidak Lulus')

# 4️. Urutkan berdasarkan nilai tertinggi
data_sorted = data.sort_values(by='Rata-rata', ascending=False)

# 5️. Tampilkan 5 mahasiswa dengan nilai tertinggi
print("=== 5 Mahasiswa dengan Nilai Tertinggi ===")
print(data_sorted.head(5))

# 6️. Simpan hasil ke file Excel baru
output_file = "rekap_nilai.xlsx"
data_sorted.to_excel(output_file, index=False)

print(f"\nFile hasil telah disimpan ke: {output_file}")