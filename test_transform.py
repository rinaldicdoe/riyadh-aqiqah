import pandas as pd
import io

# Minimal test data dengan duplikasi
test_data = """Tanggal Kirim,No. Invoice,Paket & Menu,Jumlah,Satuan,Cabang
2025-11-01,INV001,Paket Aqiqah Domba Tipe A - Betina,1,EKOR,Cibubur
2025-11-01,INV002,Paket Aqiqah Domba Tipe A - Betina,2,EKOR,Cibubur
2025-11-02,INV003,Paket Aqiqah Domba Tipe A - Betina,1,EKOR,Cibubur
2025-11-02,INV004,Menu Item 1,5,PORSI,Cibubur
2025-11-02,INV005,Menu Item 1,3,PORSI,Cibubur
"""

df = pd.read_csv(io.StringIO(test_data))
print("=== INPUT DATA ===")
print(df)
print("\n")

# Proses: split EKOR dan non-EKOR
df_ekor = df[df['Satuan'] == 'EKOR'].copy()
df_non_ekor = df[df['Satuan'] != 'EKOR'].copy()

# Process EKOR
df_ekor['Paket & Menu'] = df_ekor['Paket & Menu'].str.replace('Paket Aqiqah ', '', regex=False)
df_ekor = df_ekor.drop(columns=["No. Invoice", "Satuan"])
df_grouped_ekor = df_ekor.groupby(["Tanggal Kirim", "Paket & Menu"]).agg(Jumlah=('Jumlah', 'sum')).reset_index()

print("=== EKOR GROUPED ===")
print(df_grouped_ekor)
print("\n")

# Process non-EKOR
df_non_ekor = df_non_ekor.drop(columns=["No. Invoice", "Satuan"])
df_grouped_non_ekor = df_non_ekor.groupby(["Tanggal Kirim", "Paket & Menu"]).agg(Jumlah=('Jumlah', 'sum')).reset_index()

print("=== NON-EKOR GROUPED ===")
print(df_grouped_non_ekor)
print("\n")

# Combine
df_grouped = pd.concat([df_grouped_ekor, df_grouped_non_ekor])
print("=== COMBINED GROUPED ===")
print(df_grouped)
print("\n")

# Format dates
df_grouped['Tanggal Kirim'] = pd.to_datetime(df_grouped['Tanggal Kirim']).dt.strftime('%d-%m-%Y')

# Pivot
df_pivot = df_grouped.pivot_table(
    index='Paket & Menu',
    columns='Tanggal Kirim',
    values='Jumlah',
    aggfunc='sum'
).fillna(0)

print("=== PIVOT TABLE (BEFORE RENAME) ===")
print(df_pivot)
print("Columns:", df_pivot.columns.tolist())
print("\n")

df_pivot.reset_index(inplace=True)

# Now rename columns with weekday
pivot_other_cols = [c for c in df_pivot.columns if c != 'Paket & Menu']
print("Date columns to rename:", pivot_other_cols)

hari_id = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
renamed_map_pivot = {}
for orig in pivot_other_cols:
    try:
        dt = pd.to_datetime(orig, format='%d-%m-%Y')
        hari = hari_id[dt.weekday()]
        renamed_map_pivot[orig] = f"{hari}: {dt.strftime('%d-%m-%Y')}"
        print(f"  {orig} -> {renamed_map_pivot[orig]}")
    except Exception as e:
        print(f"  ERROR parsing {orig}: {e}")

print("\nRenamed map:", renamed_map_pivot)
print("\n")

if renamed_map_pivot:
    df_pivot = df_pivot.rename(columns=renamed_map_pivot)

print("=== PIVOT TABLE (AFTER RENAME) ===")
print(df_pivot)
print("Columns:", df_pivot.columns.tolist())
