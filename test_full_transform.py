import pandas as pd
import io

# Kategori data (simplified)
kategori_csv_data = """Nama Barang;Kelompok 
Paket Aqiqah Domba Tipe A - Betina;Paket Aqiqah Domba Tipe A - Betina
"""

# Test data
test_data = """Tanggal Kirim,No. Invoice,Paket & Menu,Jumlah,Satuan,Cabang
2025-11-01,INV001,Paket Aqiqah Domba Tipe A - Betina,1,EKOR,Cibubur
2025-11-01,INV002,Paket Aqiqah Domba Tipe A - Betina,2,EKOR,Cibubur
2025-11-02,INV003,Paket Aqiqah Domba Tipe A - Betina,1,EKOR,Cibubur
"""

df_sales = pd.read_csv(io.StringIO(test_data))
df_kategori = pd.read_csv(io.StringIO(kategori_csv_data), sep=';')
df_kategori.rename(columns={'Kelompok ': 'Paket & Menu Kategori'}, inplace=True)
df_kategori = df_kategori.drop_duplicates(subset=['Nama Barang'], keep='first')

print("=== KATEGORI ===")
print(df_kategori)
print("\n")

# Process
df_sales['Tanggal Kirim'] = pd.to_datetime(df_sales['Tanggal Kirim'], errors='coerce')
df_ekor = df_sales[df_sales['Satuan'] == 'EKOR'].copy()
df_ekor['Paket & Menu'] = df_ekor['Paket & Menu'].str.replace('Paket Aqiqah ', '', regex=False)
df_ekor = df_ekor.drop(columns=["No. Invoice", "Satuan"])
df_grouped_ekor = df_ekor.groupby(["Tanggal Kirim", "Paket & Menu"]).agg(Jumlah=('Jumlah', 'sum')).reset_index()

print("=== GROUPED EKOR ===")
print(df_grouped_ekor)
print("\n")

# Format dates and pivot
df_grouped_ekor['Tanggal Kirim'] = pd.to_datetime(df_grouped_ekor['Tanggal Kirim']).dt.strftime('%d-%m-%Y')
df_pivot = df_grouped_ekor.pivot_table(
    index='Paket & Menu',
    columns='Tanggal Kirim',
    values='Jumlah',
    aggfunc='sum'
).fillna(0)

print("=== PIVOT (before rename) ===")
print(df_pivot)
print("\n")

df_pivot.reset_index(inplace=True)

# Rename with weekday
pivot_other_cols = [c for c in df_pivot.columns if c != 'Paket & Menu']
hari_id = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
renamed_map_pivot = {}
for orig in pivot_other_cols:
    dt = pd.to_datetime(orig, format='%d-%m-%Y')
    hari = hari_id[dt.weekday()]
    renamed_map_pivot[orig] = f"{hari}: {orig}"

df_pivot = df_pivot.rename(columns=renamed_map_pivot)
new_date_cols = list(renamed_map_pivot.values())

print("=== PIVOT (after rename) ===")
print(df_pivot)
print("Columns:", df_pivot.columns.tolist())
print("\n")

# Add Kelompok
df_pivot['Kelompok'] = 'Paket Aqiqah ' + df_pivot['Paket & Menu']

print("=== PIVOT (with Kelompok) ===")
print(df_pivot)
print("\n")

# MERGE
print("=== BEFORE MERGE ===")
print(f"df_pivot shape: {df_pivot.shape}")
print(f"df_kategori shape: {df_kategori.shape}")
print("\n")

df_merged = pd.merge(df_pivot, df_kategori[['Nama Barang', 'Paket & Menu Kategori']], 
                     left_on='Kelompok', right_on='Nama Barang', how='left')

print("=== AFTER MERGE ===")
print(f"df_merged shape: {df_merged.shape}")
print(df_merged)
print("Columns:", df_merged.columns.tolist())
print("\n")

# Rename
df_merged.rename(columns={'Paket & Menu Kategori': 'Paket & Menu Final'}, inplace=True)
all_cols = list(df_merged.columns)
if 'Paket & Menu Final' in all_cols:
    all_cols.remove('Paket & Menu Final')
    all_cols.insert(0, 'Paket & Menu Final')

df_final = df_merged[all_cols]

print("=== AFTER REORDER ===")
print(df_final)
print("\n")

# Drop columns
cols_to_drop_final = ['Paket & Menu', 'Kelompok', 'Nama Barang']
df_final = df_final.drop(columns=cols_to_drop_final, errors='ignore')
df_final.rename(columns={'Paket & Menu Final': 'Paket & Menu'}, inplace=True)

print("=== AFTER DROP & RENAME ===")
print(df_final)
print("Duplicates in Paket & Menu:", df_final['Paket & Menu'].duplicated().sum())
print("\n")

df_final.dropna(subset=['Paket & Menu'], inplace=True)

print("=== FINAL ===")
print(df_final)
print("Shape:", df_final.shape)
print("Duplicates:", df_final['Paket & Menu'].duplicated().sum())
