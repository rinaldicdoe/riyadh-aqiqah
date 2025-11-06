import streamlit as st
import pandas as pd
import io
import openpyxl 
import datetime
import os
# Import library untuk Word
from docx import Document
from docx.shared import Cm, Pt, RGBColor

# Initialize session state
if 'missing_items' not in st.session_state:
    st.session_state.missing_items = None

# ============================================================
# HELPER FUNCTIONS: KATEGORI MANAGEMENT
# ============================================================
KATEGORI_FILE = "kategori.csv"

def load_kategori():
    """Load kategori from CSV file, return empty df if not exists"""
    if os.path.exists(KATEGORI_FILE):
        return pd.read_csv(KATEGORI_FILE)
    return pd.DataFrame(columns=['Nama Barang', 'Kategori Final'])

def save_kategori(df):
    """Save kategori to CSV file"""
    df.to_csv(KATEGORI_FILE, index=False)
    st.success("✅ Kategori berhasil disimpan!")

def add_kategori(nama_barang, kategori_final):
    """Add new kategori entry"""
    df_kat = load_kategori()
    new_row = pd.DataFrame({'Nama Barang': [nama_barang], 'Kategori Final': [kategori_final]})
    df_kat = pd.concat([df_kat, new_row], ignore_index=True)
    save_kategori(df_kat)
    return df_kat

def delete_kategori(idx):
    """Delete kategori by index"""
    df_kat = load_kategori()
    df_kat = df_kat.drop(idx).reset_index(drop=True)
    save_kategori(df_kat)
    return df_kat

def update_kategori(idx, nama_barang, kategori_final):
    """Update kategori entry"""
    df_kat = load_kategori()
    df_kat.at[idx, 'Nama Barang'] = nama_barang
    df_kat.at[idx, 'Kategori Final'] = kategori_final
    save_kategori(df_kat)
    return df_kat

# =============================================================================
# FUNGSI 1: TRANSFORMASI REKAP PEMOTONGAN
# =============================================================================
def transform_rekap_pemotongan(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, sheet_name="Status Pesanan Penjualan", header=1, engine='openpyxl')
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=1)
        else:
            st.error("Format file tidak didukung. Harap unggah file .xlsx atau .csv")
            return None
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.warning("Tips: Jika file .xlsx Anda error, coba simpan sheet 'Status Pesanan Penjualan' sebagai file CSV dan unggah kembali file CSV tersebut.")
        return None

    df.dropna(subset=['Cabang'], inplace=True)
    df.dropna(subset=['Tanggal Kirim'], inplace=True)
    df = df[df['Cabang'] != 'Cabang'].copy()
    
    if "Jenis Kelamin AnakNama" in df.columns:
        df.rename(columns={"Jenis Kelamin AnakNama": "Jenis Kelamin Anak"}, inplace=True)
    if "Pemotongan DisaksikanNama" in df.columns:
        df.rename(columns={"Pemotongan DisaksikanNama": "Pemotongan Disaksikan"}, inplace=True)
        
    df['Tanggal Kirim'] = pd.to_datetime(df['Tanggal Kirim'], errors='coerce').dt.date
    df['Tanggal Potong'] = pd.to_datetime(df['Tanggal Potong'], errors='coerce').dt.date
    
    for col in ['Telpon 1', 'Telpon 2']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True)

    df_menu_prep = df[['No. Invoice', 'Paket & Menu', 'Jumlah']].copy()
    df_menu_prep = df_menu_prep[~df_menu_prep['Paket & Menu'].astype(str).str.contains("Paket", na=False)]
    df_menu_prep['Menu_Item'] = df_menu_prep['Paket & Menu'].astype(str) + " " + df_menu_prep['Jumlah'].astype(str)
    df_menu_final = df_menu_prep.groupby('No. Invoice')['Menu_Item'].apply(lambda x: ', '.join(x) + ' PORSI').reset_index()
    df_menu_final.rename(columns={'Menu_Item': 'Menu'}, inplace=True)

    df_paket_prep = df[['No. Invoice', 'Paket & Menu', 'Jumlah']].copy()
    df_paket_final = df_paket_prep[df_paket_prep['Paket & Menu'].astype(str).str.contains("Paket", na=False)].copy()
    df_paket_final.drop_duplicates(subset=['No. Invoice'], keep='first', inplace=True)
    df_paket_final.rename(columns={'Paket & Menu': 'Paket'}, inplace=True)
    
    cols_to_drop = ['Paket & Menu', 'No. Urut', 'No. Domba', 'Satuan',
                    'Tanggal Domba Dipotong', 'Jam Tiba (hh:mm)', 'Jam Kirim (hh:mm)', 'Kode Menu']
    df_base = df.drop(columns=cols_to_drop + ['Jumlah'], errors='ignore')

    df_merged = pd.merge(df_base, df_paket_final[['No. Invoice', 'Paket', 'Jumlah']], on='No. Invoice', how='left')
    df_merged = pd.merge(df_merged, df_menu_final[['No. Invoice', 'Menu']], on='No. Invoice', how='left')

    df_final = df_merged.drop_duplicates(subset=['No. Invoice'], keep='first').copy()
    df_final['Pemotongan Real'] = ''

    final_columns_order = [
        "Cabang", "Tanggal Kirim", "Tanggal Potong", "No. Invoice", "Status Perkembangan", 
        "Pemotongan Real", "Nama Anak", "Jenis Kelamin Anak", "Nama Bapak", 
        "Telpon 1", "Telpon 2", "Paket", "Jumlah", "Menu", "Pemotongan Disaksikan", 
        "Catatan Khusus", "CS"
    ]
    
    existing_columns = [col for col in final_columns_order if col in df_final.columns]
    df_final = df_final[existing_columns]
    df_final.insert(0, 'Nomor', range(1, len(df_final) + 1))
    
    return df_final

# =============================================================================
# FUNGSI 2: TRANSFORMASI REKAP KEBUTUHAN MINGGUAN
# =============================================================================
def transform_rekap_kebutuhan(file_sales):
    try:
        df_sales = pd.read_excel(file_sales, sheet_name="Status Pesanan Penjualan", header=1, engine='openpyxl')
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.warning("Pastikan nama sheet pada file sales adalah 'Status Pesanan Penjualan'.")
        return None

    # Load kategori from external file
    df_kategori = load_kategori()
    
    # Clean up data
    df_sales.dropna(subset=['Tanggal Kirim', 'No. Invoice'], inplace=True)
    df_sales = df_sales[df_sales['Cabang'] != 'Cabang'].copy()
    
    # Parse dates
    df_sales['Tanggal Kirim'] = pd.to_datetime(df_sales['Tanggal Kirim'], errors='coerce')
    
    # Get Paket & Menu from column Q (index 16) and Jumlah from column R (index 17)
    try:
        # Get the column names (Excel might have them or might not)
        if 'Paket & Menu' not in df_sales.columns:
            # Try by column index
            paket_col = df_sales.iloc[:, 16]
            jumlah_col = df_sales.iloc[:, 17]
        else:
            paket_col = df_sales['Paket & Menu']
            jumlah_col = df_sales['Jumlah']
        
        df_sales['Paket & Menu'] = paket_col
        df_sales['Jumlah'] = pd.to_numeric(jumlah_col, errors='coerce').fillna(0).astype(int)
    except Exception as e:
        st.error(f"Tidak dapat menemukan data paket & menu atau jumlah: {e}")
        return None
    
    # Groupby Tanggal Kirim and Paket & Menu, aggregating Jumlah
    df_grouped = df_sales.groupby(["Tanggal Kirim", "Paket & Menu"])['Jumlah'].sum().reset_index()
    df_grouped['Jumlah'] = df_grouped['Jumlah'].astype(int)
    
    # Format dates to string for pivot table
    df_grouped['Tanggal Kirim_str'] = df_grouped['Tanggal Kirim'].apply(
        lambda x: pd.to_datetime(x).strftime('%d-%m-%Y')
    )
    
    # Create pivot table
    df_pivot = df_grouped.pivot_table(
        index='Paket & Menu',
        columns='Tanggal Kirim_str',
        values='Jumlah',
        aggfunc='sum'
    ).fillna(0).astype(int)
    
    df_pivot.reset_index(inplace=True)
    
    # Sort date columns
    date_cols = [c for c in df_pivot.columns if c != 'Paket & Menu']
    date_cols_sorted = sorted(date_cols, key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
    
    # Add weekday names to date columns
    hari_id = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
    date_cols_renamed = {}
    for date_str in date_cols_sorted:
        date_obj = pd.to_datetime(date_str, format='%d-%m-%Y')
        weekday_name = hari_id[date_obj.weekday()]
        new_col_name = f"{weekday_name}: {date_str}"
        date_cols_renamed[date_str] = new_col_name
    
    df_pivot.rename(columns=date_cols_renamed, inplace=True)
    date_cols_sorted_renamed = [date_cols_renamed[c] for c in date_cols_sorted]
    
    # Reorder columns
    final_col_order = ['Paket & Menu'] + date_cols_sorted_renamed
    df_final = df_pivot[final_col_order].copy()
    
    # Check if items are in kategori and store missing items
    st.session_state.missing_items = None
    if len(df_kategori) > 0:
        missing_items_list = []
        for item in df_final['Paket & Menu']:
            if item not in df_kategori['Nama Barang'].values:
                missing_items_list.append(item)
        
        if len(missing_items_list) > 0:
            df_missing = df_final[df_final['Paket & Menu'].isin(missing_items_list)].copy()
            st.session_state.missing_items = df_missing
    
    # Add TOTAL row at the bottom
    if len(date_cols_sorted_renamed) > 0:
        totals_row = {'Paket & Menu': 'TOTAL'}
        for col in date_cols_sorted_renamed:
            totals_row[col] = int(df_final[col].sum())
        df_final = pd.concat([df_final, pd.DataFrame([totals_row])], ignore_index=True)
    
    return df_final

# =============================================================================
# FUNGSI 3: TRANSFORMASI LABEL MASAK
# =============================================================================
def transform_and_create_word_label(file_input):
    try:
        df = pd.read_excel(file_input, sheet_name="Status Pesanan Penjualan", header=1)
        df.dropna(subset=['No. Invoice'], inplace=True)
        df = df[df['Cabang'] != 'Cabang'].copy()
        
        for col in ['Tanggal Kirim', 'Tanggal Potong']:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')
        
        def format_time(time_val):
            if isinstance(time_val, datetime.time):
                return time_val.strftime('%H:%M')
            elif isinstance(time_val, str):
                return time_val
            return str(time_val)

        df['Jam Tiba (hh:mm)'] = df['Jam Tiba (hh:mm)'].apply(format_time)
        df['Jam Kirim (hh:mm)'] = df['Jam Kirim (hh:mm)'].apply(format_time)
        df.fillna('', inplace=True)
        
        df['Detail Customer'] = (
            "Nama Aqiqah: " + df['Nama Anak'].astype(str).str.strip() + "\n" +
            "No. Invoice: " + df['No. Invoice'].astype(str).str.strip() + "\n" +
            "Jenis Kelamin: " + df['Jenis Kelamin Anak'].astype(str).str.strip() + "\n" +
            "Cabang: " + df['Cabang'].astype(str).str.strip()
        )
        df['Detail Waktu'] = (
            "Tgl Kirim: " + df['Tanggal Kirim'].astype(str).str.strip() + "\n" +
            "Tgl Potong: " + df['Tanggal Potong'].astype(str).str.strip() + "\n" +
            "Jam Tiba: " + df['Jam Tiba (hh:mm)'].astype(str).str.strip() + "\n" +
            "Jam Kirim: " + df['Jam Kirim (hh:mm)'].astype(str).str.strip()
        )
        df['Menu'] = df['Paket & Menu'].astype(str) + " " + df['Jumlah'].astype(str) + " " + df['Satuan'].astype(str)
        df['Berat'] = "Berat |\n....... KG"
        df_final = df[['Detail Customer', 'Detail Waktu', 'Menu', 'Berat']].copy()
        
        doc = Document()
        for section in doc.sections:
            section.top_margin = Cm(0.88)
            section.bottom_margin = Cm(1.75)
            section.left_margin = Cm(2.12)
            section.right_margin = Cm(1.42)

        for i in range(0, len(df_final), 5):
            chunk = df_final.iloc[i:i+5]
            table = doc.add_table(rows=0, cols=4)
            table.style = 'Table Grid'
            
            for _, record in chunk.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(record['Detail Customer'])
                row_cells[1].text = str(record['Detail Waktu'])
                row_cells[2].text = str(record['Menu'])
                row_cells[3].text = str(record['Berat'])
            
            for row in table.rows:
                row.height = Cm(4.5)
                for cell in row.cells:
                    text_to_check = cell.text
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.line_spacing = 1.3
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
                            if "Cibubur" in text_to_check:
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(0, 0, 255)
                            elif "JaDeTa" in text_to_check:
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(255, 0, 0)
            
            table.columns[0].width = Cm(5.64)
            table.columns[1].width = Cm(2.75)
            table.columns[2].width = Cm(5.29)
            table.columns[3].width = Cm(1.76)
            
            if i + 5 < len(df_final):
                doc.add_page_break()
        
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        return doc_io
    except Exception as e:
        st.error(f"Terjadi kesalahan pada pembuatan Label Masak: {e}")
        import traceback
        st.text(traceback.format_exc())
        return None

# =============================================================================
# BAGIAN TAMPILAN APLIKASI WEB (USER INTERFACE)
# =============================================================================
st.sidebar.title("Navigasi Menu")
menu_pilihan = st.sidebar.radio(
    "Pilih menu yang ingin Anda gunakan:",
    ("Rekap Pemotongan", "Rekap Kebutuhan Mingguan", "Label Masak")
)

# --- TAMPILAN MENU 1: REKAP PEMOTONGAN ---
if menu_pilihan == "Rekap Pemotongan":
    st.title("📝 Rekap Pemotongan")
    st.write("Unggah file Excel mentah untuk diproses.")
    uploaded_file_rekap = st.file_uploader("Pilih file Excel atau CSV", type=['xlsx', 'csv'], key="rekap_pemotongan")
    if uploaded_file_rekap:
        st.info(f"✔️ File diterima: **{uploaded_file_rekap.name}**. Memproses...")
        result_df_rekap = transform_rekap_pemotongan(uploaded_file_rekap)
        if result_df_rekap is not None:
            st.success("🎉 Transformasi data berhasil!")
            st.dataframe(result_df_rekap)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df_rekap.to_excel(writer, index=False, sheet_name='Hasil Rekap Pemotongan')
                workbook = writer.book
                worksheet = writer.sheets['Hasil Rekap Pemotongan']
                general_format = workbook.add_format({'font_size': 22, 'valign': 'vcenter', 'align': 'center', 'text_wrap': True})
                header_format = workbook.add_format({'font_size': 22, 'bold': True, 'valign': 'vcenter', 'align': 'center', 'text_wrap': True})
                worksheet.set_margins(left=0.28, right=0.28, top=0.75, bottom=0.75)
                worksheet.set_landscape()
                worksheet.set_default_row(257)
                worksheet.set_row(0, 60, header_format)
                worksheet.hide_gridlines(2)
                worksheet.fit_to_pages(1, 0)
                worksheet.set_column('A:A', 7.40); worksheet.set_column('B:B', 15.27); worksheet.set_column('C:D', 26.07)
                worksheet.set_column('E:E', 22.47); worksheet.set_column('F:F', 13.27); worksheet.set_column('G:G', 20)
                worksheet.set_column('H:L', 22.73); worksheet.set_column('M:M', 32.47); worksheet.set_column('N:N', 15)
                worksheet.set_column('O:O', 35); worksheet.set_column('P:P', 25); worksheet.set_column('Q:Q', 45)
                worksheet.set_column('R:R', 7.40)
                (max_row, max_col) = result_df_rekap.shape
                worksheet.set_column(0, max_col - 1, None, general_format)
            excel_data = output.getvalue()
            now = datetime.datetime.now()
            download_filename = now.strftime("%d_%m_%Y-%H_%M") + "-Rekap_Pemotongan.xlsx"
            st.download_button(label="⬇️ Download Hasil sebagai Excel", data=excel_data, file_name=download_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- TAMPILAN MENU 2: REKAP KEBUTUHAN MINGGUAN ---
elif menu_pilihan == "Rekap Kebutuhan Mingguan":
    st.title("📊 Rekap Kebutuhan Mingguan")
    st.write("Unggah **File Status Penjualan** untuk diproses.")
    
    # ===== KATEGORI MANAGEMENT (EXPANDER) =====
    with st.expander("⚙️ Kelola Kategori Paket & Menu", expanded=False):
        st.subheader("Kategori Saat Ini")
        df_kategori = load_kategori()
        
        if len(df_kategori) > 0:
            st.dataframe(df_kategori, use_container_width=True, height=250)
        else:
            st.info("Belum ada kategori.")
        
        # Tab untuk operasi kategori
        kat_tab1, kat_tab2, kat_tab3 = st.tabs(["➕ Tambah", "✏️ Edit", "🗑️ Hapus"])
        
        # TAB 1: Tambah Kategori
        with kat_tab1:
            col1, col2 = st.columns(2)
            with col1:
                nama_barang_baru = st.text_input("Nama Barang / Paket & Menu:", key="nama_barang_tambah_kat")
            with col2:
                kategori_final_baru = st.text_input("Kategori Final:", key="kategori_final_tambah_kat")
            
            if st.button("✅ Tambah", key="btn_tambah_kat"):
                if nama_barang_baru.strip() and kategori_final_baru.strip():
                    df_kat_check = load_kategori()
                    if nama_barang_baru in df_kat_check['Nama Barang'].values:
                        st.warning("⚠️ Nama barang sudah ada.")
                    else:
                        add_kategori(nama_barang_baru, kategori_final_baru)
                        st.rerun()
                else:
                    st.error("❌ Tidak boleh kosong!")
        
        # TAB 2: Edit Kategori
        with kat_tab2:
            df_kat_edit = load_kategori()
            if len(df_kat_edit) > 0:
                idx_edit = st.selectbox("Pilih kategori untuk diedit:", 
                                       range(len(df_kat_edit)), 
                                       format_func=lambda i: f"{i}. {df_kat_edit.iloc[i]['Nama Barang']}",
                                       key="select_edit_kat")
                
                current_nama = df_kat_edit.iloc[idx_edit]['Nama Barang']
                current_kategori = df_kat_edit.iloc[idx_edit]['Kategori Final']
                
                col1, col2 = st.columns(2)
                with col1:
                    nama_barang_edit = st.text_input("Nama Barang:", value=current_nama, key="nama_barang_edit_kat")
                with col2:
                    kategori_final_edit = st.text_input("Kategori Final:", value=current_kategori, key="kategori_final_edit_kat")
                
                if st.button("💾 Simpan", key="btn_edit_kat"):
                    if nama_barang_edit.strip() and kategori_final_edit.strip():
                        update_kategori(idx_edit, nama_barang_edit, kategori_final_edit)
                        st.rerun()
                    else:
                        st.error("❌ Tidak boleh kosong!")
            else:
                st.info("Tidak ada kategori untuk diedit.")
        
        # TAB 3: Hapus Kategori
        with kat_tab3:
            df_kat_delete = load_kategori()
            if len(df_kat_delete) > 0:
                idx_hapus = st.selectbox("Pilih kategori untuk dihapus:", 
                                        range(len(df_kat_delete)), 
                                        format_func=lambda i: f"{i}. {df_kat_delete.iloc[i]['Nama Barang']}",
                                        key="select_delete_kat")
                
                st.warning(f"⚠️ Akan dihapus: **{df_kat_delete.iloc[idx_hapus]['Nama Barang']}**")
                
                if st.button("🗑️ Hapus", key="btn_hapus_kat"):
                    delete_kategori(idx_hapus)
                    st.rerun()
            else:
                st.info("Tidak ada kategori untuk dihapus.")
        
        # Backup & Restore
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            df_kat_backup = load_kategori()
            if st.button("⬇️ Download kategori.csv"):
                csv_data = df_kat_backup.to_csv(index=False)
                st.download_button(
                    label="📥 Download",
                    data=csv_data,
                    file_name="kategori_backup.csv",
                    mime="text/csv"
                )
        
        with col2:
            uploaded_kategori = st.file_uploader("📤 Upload kategori.csv:", type=['csv'], key="upload_kategori_kat")
            if uploaded_kategori is not None:
                try:
                    df_upload = pd.read_csv(uploaded_kategori)
                    if 'Nama Barang' in df_upload.columns and 'Kategori Final' in df_upload.columns:
                        save_kategori(df_upload)
                        st.rerun()
                    else:
                        st.error("❌ Format file salah!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")
    
    st.divider()
    
    # ===== MAIN PROCESSING =====
    uploaded_file_sales = st.file_uploader("Pilih File Excel Penjualan", type=['xlsx'], key="status_penjualan")
    if uploaded_file_sales:
        st.info(f"✔️ File '{uploaded_file_sales.name}' diterima. Memproses...")
        result_df_kebutuhan = transform_rekap_kebutuhan(uploaded_file_sales)
        if result_df_kebutuhan is not None:
            st.success("🎉 Transformasi data berhasil!")
            st.dataframe(result_df_kebutuhan)
            
            # --- Display items not in kategori ---
            if 'missing_items' in st.session_state and st.session_state.missing_items is not None:
                df_missing = st.session_state.missing_items
                if len(df_missing) > 0:
                    st.warning("⚠️ **Item yang tidak ada di Kategori:**")
                    
                    # Prepare display with total from last date column (R)
                    display_cols = df_missing.columns.tolist()
                    if len(display_cols) > 1:
                        # Get last date column (R column)
                        last_date_col = display_cols[-1]
                        
                        # Create summary df: Paket & Menu with total from last column
                        df_display = df_missing[['Paket & Menu', last_date_col]].copy()
                        df_display.rename(columns={last_date_col: 'Total (Kolom R)'}, inplace=True)
                        df_display['Total (Kolom R)'] = df_display['Total (Kolom R)'].astype(int)
                        df_display.rename(columns={'Paket & Menu': 'Nama Paket & Menu'}, inplace=True)
                        
                        st.dataframe(df_display, use_container_width=True)
            
            output_kebutuhan = io.BytesIO()
            with pd.ExcelWriter(output_kebutuhan, engine='xlsxwriter') as writer:
                result_df_kebutuhan.to_excel(writer, index=False, sheet_name='Rekap Kebutuhan', startrow=1, header=False)
                workbook = writer.book
                worksheet = writer.sheets['Rekap Kebutuhan']
                header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
                for col_num, value in enumerate(result_df_kebutuhan.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(0, 0, 40)
                worksheet.set_column(1, len(result_df_kebutuhan.columns)-1, 15)
            excel_data_kebutuhan = output_kebutuhan.getvalue()
            now = datetime.datetime.now()
            download_filename_kebutuhan = now.strftime("%d_%m_%Y-%H_%M") + "-Rekap_Kebutuhan_Mingguan.xlsx"
            st.download_button(label="⬇️ Download Rekap Kebutuhan Mingguan", data=excel_data_kebutuhan, file_name=download_filename_kebutuhan, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- TAMPILAN MENU 3: LABEL MASAK ---
elif menu_pilihan == "Label Masak":
    st.title("🏷️ Label Masak")
    st.write("Unggah file template Excel untuk membuat label masak dalam format Microsoft Word.")
    uploaded_file_label = st.file_uploader("Pilih File Template Excel", type=['xlsx'], key="label_masak")
    if uploaded_file_label is not None:
        st.info(f"✔️ File '{uploaded_file_label.name}' diterima. Memproses...")
        word_file_buffer = transform_and_create_word_label(uploaded_file_label)
        if word_file_buffer:
            st.success("🎉 Dokumen Word berhasil dibuat!")
            now = datetime.datetime.now()
            download_filename_word = now.strftime("%d_%m_%Y-%H_%M") + "-Label_Masak.docx"
            st.download_button(
                label="⬇️ Download Label Masak sebagai Word",
                data=word_file_buffer,
                file_name=download_filename_word,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )