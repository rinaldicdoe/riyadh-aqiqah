import streamlit as st
import pandas as pd
import io
import openpyxl 
import datetime
# Import library untuk Word, termasuk RGBColor
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION

# =============================================================================
# FUNGSI 1: TRANSFORMASI REKAP PEMOTONGAN
# =============================================================================
def transform_rekap_pemotongan(uploaded_file):
    # ... (Kode fungsi ini tidak berubah, dibiarkan sama seperti sebelumnya)
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
    # ... (Kode fungsi ini tidak berubah, dibiarkan sama seperti sebelumnya)
    try:
        df_sales = pd.read_excel(file_sales, sheet_name="Status Pesanan Penjualan", header=1)
        
        kategori_csv_data = """No,Kode Barang,Nama Barang,Kelompok 
1,1PKT-00012,Paket Aqiqah Domba Kebuli - Betina,Paket Aqiqah Domba Kebuli - Betina
2,1PKT-00022,Paket Aqiqah Domba Kebuli - Betina (Nasi Biasa),Paket Aqiqah Domba Kebuli - Betina
3,1PKT-00013,Paket Aqiqah Domba Kebuli - Jantan,Paket Aqiqah Domba Kebuli - Jantan
4,1PKT-00023,Paket Aqiqah Domba Kebuli - Jantan (Nasi Biasa),Paket Aqiqah Domba Kebuli - Jantan
5,1PKT-00001,Paket Aqiqah Domba Tipe A - Betina,Paket Aqiqah Domba Tipe A - Betina
6,1PKT-00052,Paket Aqiqah Domba Tipe A - Betina (Bento Spesial),Paket Aqiqah Domba Tipe A - Betina
7,1PKT-00062,Paket Aqiqah Domba Tipe A - Betina (Hemat),Paket Aqiqah Domba Tipe A - Betina
8,1PKT-00014,Paket Aqiqah Domba Tipe A - Betina (Spesial),Paket Aqiqah Domba Tipe A - Betina
9,1PKT-00032,Paket Aqiqah Domba Tipe A - Betina (Tanpa Masak),Paket Aqiqah Domba Tipe A - Betina
10,1PKT-00024,Paket Aqiqah Domba Tipe A - Betina (Tanpa Nasi Box),Paket Aqiqah Domba Tipe A - Betina
11,1PKT-00005,Paket Aqiqah Domba Tipe A - Jantan,Paket Aqiqah Domba Tipe A - Jantan
12,1PKT-00049,Paket Aqiqah Domba Tipe A - Jantan (Plus),Paket Aqiqah Domba Tipe A - Jantan
13,1PKT-00018,Paket Aqiqah Domba Tipe A - Jantan (Spesial),Paket Aqiqah Domba Tipe A - Jantan
14,1PKT-00036,Paket Aqiqah Domba Tipe A - Jantan (Tanpa Masak),Paket Aqiqah Domba Tipe A - Jantan
15,1PKT-00028,Paket Aqiqah Domba Tipe A - Jantan (Tanpa Nasi Box),Paket Aqiqah Domba Tipe A - Jantan
16,1PKT-00002,Paket Aqiqah Domba Tipe B - Betina,Paket Aqiqah Domba Tipe B - Betina
17,1PKT-00015,Paket Aqiqah Domba Tipe B - Betina (Spesial),Paket Aqiqah Domba Tipe B - Betina
18,1PKT-00033,Paket Aqiqah Domba Tipe B - Betina (Tanpa Masak),Paket Aqiqah Domba Tipe B - Betina
19,1PKT-00025,Paket Aqiqah Domba Tipe B - Betina (Tanpa Nasi Box),Paket Aqiqah Domba Tipe B - Betina
20,1PKT-00006,Paket Aqiqah Domba Tipe B - Jantan,Paket Aqiqah Domba Tipe B - Jantan
21,1PKT-00019,Paket Aqiqah Domba Tipe B - Jantan (Spesial),Paket Aqiqah Domba Tipe B - Jantan
22,1PKT-00037,Paket Aqiqah Domba Tipe B - Jantan (Tanpa Masak),Paket Aqiqah Domba Tipe B - Jantan
23,1PKT-00029,Paket Aqiqah Domba Tipe B - Jantan (Tanpa Nasi Box),Paket Aqiqah Domba Tipe B - Jantan
24,1PKT-00003,Paket Aqiqah Domba Tipe C - Betina,Paket Aqiqah Domba Tipe C - Betina
25,1PKT-00016,Paket Aqiqah Domba Tipe C - Betina (Spesial),Paket Aqiqah Domba Tipe C - Betina
26,1PKT-00034,Paket Aqiqah Domba Tipe C - Betina (Tanpa Masak),Paket Aqiqah Domba Tipe C - Betina
27,1PKT-00026,Paket Aqiqah Domba Tipe C - Betina (Tanpa Nasi Box),Paket Aqiqah Domba Tipe C - Betina
28,1PKT-00007,Paket Aqiqah Domba Tipe C - Jantan,Paket Aqiqah Domba Tipe C - Jantan
29,1PKT-00020,Paket Aqiqah Domba Tipe C - Jantan (Spesial),Paket Aqiqah Domba Tipe C - Jantan
30,1PKT-00038,Paket Aqiqah Domba Tipe C - Jantan (Tanpa Masak),Paket Aqiqah Domba Tipe C - Jantan
31,1PKT-00030,Paket Aqiqah Domba Tipe C - Jantan (Tanpa Nasi Box),Paket Aqiqah Domba Tipe C - Jantan
32,1PKT-00004,Paket Aqiqah Domba Tipe D - Betina,Paket Aqiqah Domba Tipe D - Betina
33,1PKT-00017,Paket Aqiqah Domba Tipe D - Betina (Spesial),Paket Aqiqah Domba Tipe D - Betina
34,1PKT-00035,Paket Aqiqah Domba Tipe D - Betina (Tanpa Masak),Paket Aqiqah Domba Tipe D - Betina
35,1PKT-00027,Paket Aqiqah Domba Tipe D - Betina (Tanpa Nasi Box),Paket Aqiqah Domba Tipe D - Betina
36,1PKT-00008,Paket Aqiqah Domba Tipe D - Jantan,Paket Aqiqah Domba Tipe D - Jantan
37,1PKT-00021,Paket Aqiqah Domba Tipe D - Jantan (Spesial),Paket Aqiqah Domba Tipe D - Jantan
38,1PKT-00039,Paket Aqiqah Domba Tipe D - Jantan (Tanpa Masak),Paket Aqiqah Domba Tipe D - Jantan
39,1PKT-00031,Paket Aqiqah Domba Tipe D - Jantan (Tanpa Nasi Box),Paket Aqiqah Domba Tipe D - Jantan
        """
        df_kategori = pd.read_csv(io.StringIO(kategori_csv_data))

    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.warning("Pastikan nama sheet pada file sales adalah 'Status Pesanan Penjualan'.")
        return None
    
    df_kategori.dropna(subset=['Kode Barang'], inplace=True)
    if 'Kelompok ' in df_kategori.columns:
        df_kategori.rename(columns={'Kelompok ': 'Paket & Menu Kategori'}, inplace=True)

    df_sales['Tanggal Potong'] = pd.to_datetime(df_sales['Tanggal Potong'], errors='coerce')
    df_sales.dropna(subset=['Tanggal Potong', 'No. Invoice'], inplace=True)
    df_sales = df_sales[df_sales['Cabang'] != 'Cabang'].copy()
    cols_to_keep = ["Tanggal Potong", "No. Invoice", "Paket & Menu", "Jumlah", "Satuan"]
    df_sales = df_sales[cols_to_keep]
    df_sales = df_sales[df_sales['Satuan'] == 'EKOR'].copy()
    df_sales['Paket & Menu'] = df_sales['Paket & Menu'].str.replace('Paket Aqiqah ', '', regex=False)
    df_sales = df_sales.drop(columns=["No. Invoice", "Satuan"])
    df_grouped = df_sales.groupby(["Tanggal Potong", "Paket & Menu"]).agg(Jumlah=('Jumlah', 'sum')).reset_index()
    
    df_grouped['Tanggal Potong'] = pd.to_datetime(df_grouped['Tanggal Potong']).dt.strftime('%Y-%m-%d')
    df_pivot = df_grouped.pivot_table(index='Paket & Menu', columns='Tanggal Potong', values='Jumlah', aggfunc='sum').fillna(0)
    
    df_pivot.reset_index(inplace=True)
    df_pivot['Kelompok'] = 'Paket Aqiqah ' + df_pivot['Paket & Menu']
    
    df_merged = pd.merge(df_pivot, df_kategori[['Nama Barang', 'Paket & Menu Kategori']], left_on='Kelompok', right_on='Nama Barang', how='left')

    if 'Paket & Menu Kategori' in df_merged.columns:
        df_merged.rename(columns={'Paket & Menu Kategori': 'Paket & Menu Final'}, inplace=True)
        all_cols = list(df_merged.columns)
        all_cols.remove('Paket & Menu Final')
        all_cols.insert(0, 'Paket & Menu Final')
        df_final = df_merged[all_cols]
    else:
        df_final = df_merged.copy()

    cols_to_drop_final = ['Paket & Menu', 'Kelompok', 'Nama Barang']
    df_final = df_final.drop(columns=cols_to_drop_final, errors='ignore')
    df_final.rename(columns={'Paket & Menu Final': 'Paket & Menu'}, inplace=True)
    
    df_final.dropna(subset=['Paket & Menu'], inplace=True)
    
    return df_final

# =============================================================================
# FUNGSI 3: TRANSFORMASI DAN PEMBUATAN DOKUMEN WORD LABEL MASAK
# =============================================================================
def transform_and_create_word_label(file_input):
    """
    Membaca file Excel, mentransformasikannya, dan menghasilkan file Word.
    """
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
            "Tanggal Kirim: " + df['Tanggal Kirim'].astype(str).str.strip() + "\n" +
            "Tanggal Potong: " + df['Tanggal Potong'].astype(str).str.strip() + "\n" +
            "Jam Tiba: " + df['Jam Tiba (hh:mm)'].astype(str).str.strip() + "\n" +
            "Jam Kirim: " + df['Jam Kirim (hh:mm)'].astype(str).str.strip()
        )
        df['Menu'] = df['Paket & Menu'].astype(str) + " " + df['Jumlah'].astype(str) + " " + df['Satuan'].astype(str)
        df['Berat'] = "Berat |\n....... KG"

        df_final = df[['Detail Customer', 'Detail Waktu', 'Menu', 'Berat']].copy()
        
        doc = Document()
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(0.88)
            section.bottom_margin = Cm(1.75)
            section.left_margin = Cm(2.12)
            section.right_margin = Cm(1.42)

        for i in range(0, len(df_final), 5):
            chunk = df_final.iloc[i:i+5]
            
            # <<< PERBAIKAN: Menambahkan baris header di setiap tabel >>>
            table = doc.add_table(rows=1, cols=4) # Mulai dengan 1 baris untuk header
            table.style = 'Table Grid'
            
            # Isi header
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Detail Customer'
            hdr_cells[1].text = 'Detail Waktu'
            hdr_cells[2].text = 'Menu'
            hdr_cells[3].text = 'Berat'

            # Tambahkan baris data
            for df_index, record in chunk.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(record['Detail Customer'])
                row_cells[1].text = str(record['Detail Waktu'])
                row_cells[2].text = str(record['Menu'])
                row_cells[3].text = str(record['Berat'])

            # Atur format untuk seluruh tabel (termasuk header dan data)
            for row in table.rows:
                row.height = Cm(4.5)
                for cell in row.cells:
                    text_to_check = cell.text
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
                            
                            if "Cibubur" in text_to_check:
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(0, 0, 255) # Biru
                            elif "JaDeTa" in text_to_check:
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(255, 0, 0) # Merah

            # Atur lebar kolom
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
        st.error(f"Terjadi kesalahan: {e}")
        st.warning("Pastikan file yang diunggah adalah template yang benar dan berisi sheet 'Status Pesanan Penjualan'.")
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
    st.title("🚀 Aplikasi Transformasi - Rekap Pemotongan")
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
    st.title("📊 Aplikasi Transformasi - Rekap Kebutuhan Mingguan")
    st.write("Unggah **File Status Penjualan** untuk diproses.")
    uploaded_file_sales = st.file_uploader("Pilih File Excel Penjualan", type=['xlsx'], key="status_penjualan")
    if uploaded_file_sales:
        st.info(f"✔️ File '{uploaded_file_sales.name}' diterima. Memproses...")
        result_df_kebutuhan = transform_rekap_kebutuhan(uploaded_file_sales)
        if result_df_kebutuhan is not None:
            st.success("🎉 Transformasi data berhasil!")
            st.dataframe(result_df_kebutuhan)
            output_kebutuhan = io.BytesIO()
            with pd.ExcelWriter(output_kebutuhan, engine='xlsxwriter') as writer:
                result_df_kebutuhan.to_excel(writer, index=False, sheet_name='Hasil Rekap Kebutuhan')
            excel_data_kebutuhan = output_kebutuhan.getvalue()
            now = datetime.datetime.now()
            download_filename_kebutuhan = now.strftime("%d_%m_%Y-%H_%M") + "-Rekap_Kebutuhan_Mingguan.xlsx"
            st.download_button(label="⬇️ Download Rekap Kebutuhan Mingguan", data=excel_data_kebutuhan, file_name=download_filename_kebutuhan, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- TAMPILAN MENU 3: LABEL MASAK ---
elif menu_pilihan == "Label Masak":
    st.title("🏷️ Aplikasi Transformasi - Label Masak")
    st.write("Unggah file template Excel untuk membuat label masak dalam format Microsoft Word.")

    uploaded_file_label = st.file_uploader(
        "Pilih File Template Excel",
        type=['xlsx'],
        key="label_masak"
    )

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