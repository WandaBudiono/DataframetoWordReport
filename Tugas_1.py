import streamlit as st
import openpyxl
import pandas as pd
from docx import Document
from io import BytesIO


# Fungsi untuk menghitung kolom nilai dan keterangan
def calculate_rating(row):
    if pd.notnull(row['TOT POINT']):
        if row['TOT POINT'] <= 40:
            return 'E'
        elif row['TOT POINT'] <= 60:
            return 'D'
        elif row['TOT POINT'] <= 70:
            return 'C'
        elif row['TOT POINT'] <= 89:
            return 'B'
        else:
            return 'A'
    return ''

def calculate_keterangan(row):
    if row['nilai'] == 'E':
        return 'Sangat Buruk / tidak dilanjutkan kerjasamanya'
    elif row['nilai'] == 'D':
        return 'Buruk / dipertimbangkan kerjasamanya dan diberikan surat peringatan'
    elif row['nilai'] == 'C':
        return 'Cukup / dilanjutkan kerjasamanya'
    elif row['nilai'] == 'B':
        return 'Baik / dilanjutkan kerjasamanya'
    elif row['nilai'] == 'A':
        return 'Sangat Baik / dilanjutkan kerjasamanya'
    return ''

def read_excel_and_drop_nan(file_path, sheet_name):
    # Membaca file Excel
    wb = openpyxl.load_workbook(file_path)

    try:
        # Memilih sheet berdasarkan nama
        sheet = wb[sheet_name]
    except KeyError:
        print(f"Sheet dengan nama '{sheet_name}' tidak ditemukan.")
        return None
    else:
        # Mencari baris indeks yang berisi teks "NO Vendor_name, TOT.PO, RP/PP, KUALITAS, K3, L, TOT, POINT, Nilai, KETERANGAN"
        target_row_index = None
        for row in sheet.iter_rows(min_row=1, max_row=10):
            row_values = [cell.value for cell in row]
            if "NO" in row_values and "Vendor_name" in row_values and "TOT.PO" in row_values:
                target_row_index = row[0].row
                break

        if target_row_index is None:
            print("Tidak dapat menemukan baris yang sesuai di sheet", sheet_name)
            return None
        else:
            # Membaca DataFrame dari baris yang sesuai
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=target_row_index - 1, usecols="A:J")

            def drop_nan_rows_from_top(df):
                # Mencari baris pertama dengan semua nilai NaN
                first_nan_row_index = df.apply(lambda row: row.isnull().all(), axis=1).idxmax()

                # Menghapus baris-baris dengan semua nilai NaN dari bagian atas dataframe
                df = df.loc[first_nan_row_index:].copy()

                return df

            # Menghapus baris dengan nilai NaN dari bagian atas dataframe
            df_cleaned = df.drop(drop_nan_rows_from_top(df).index)

            # Menambahkan kolom 'bulan' berdasarkan kata ke-2 dari sheet_name
            df_cleaned['BULAN'] = sheet_name.split()[1]

            # Menambahkan kolom 'tahun' berdasarkan kata ke-3 dari sheet_name
            df_cleaned['TAHUN'] = sheet_name.split()[2]

            return df_cleaned

def read_all_sheets_in_excel(file_path):
    # Membaca file Excel
    wb = openpyxl.load_workbook(file_path)

    # Menginisialisasi list untuk menyimpan DataFrame dari setiap sheet
    all_dataframes = []

    for sheet_name in wb.sheetnames:
        df = read_excel_and_drop_nan(file_path, sheet_name)
        if df is not None:
            all_dataframes.append(df)

    # Menggabungkan semua DataFrame menjadi satu DataFrame
    result_df = pd.concat(all_dataframes, ignore_index=True)

    return result_df

# Fungsi untuk menampilkan DataFrame yang telah difilter berdasarkan vendor, rentang bulan, dan tahun
def display_filtered_data(df):
    global filtered_data, summary_df, doc
    # Tampilkan widget untuk memilih vendor
    selected_vendors = st.multiselect('Select Vendor:', df['Vendor_name'].unique())

    # Filter DataFrame berdasarkan vendor yang dipilih
    filtered_data = df[df['Vendor_name'].isin(selected_vendors)] if selected_vendors else df

    # Tampilkan widget untuk memilih tahun
    selected_years = st.multiselect('Select Year:', df['TAHUN'].unique())

    # Filter DataFrame berdasarkan tahun yang dipilih
    filtered_data = filtered_data[filtered_data['TAHUN'].isin(selected_years)] if selected_years else filtered_data

    # Tampilkan beberapa pilihan bulan
    months = df['BULAN'].unique()
    selected_months = st.multiselect('Select Months:', months)

    # Filter DataFrame berdasarkan bulan yang dipilih
    filtered_data = filtered_data[filtered_data['BULAN'].isin(selected_months)] if selected_months else filtered_data

    # Tampilkan DataFrame yang sudah difilter
    st.dataframe(filtered_data)
    
    # Hitung sum dan average (mean) dari kolom yang dipilih
    if not filtered_data.empty:
        sum_selected_columns = filtered_data['TOT.PO'].sum()
        mean_selected_columns = filtered_data[['RP/PP', 'KUALITAS', 'K3', 'L', 'TOT POINT']].mean()


        summary_df = pd.DataFrame({
            'Sum of TOT.PO': [int(sum_selected_columns)],
            'Mean of RP/PP': [int(mean_selected_columns['RP/PP'])],
            'Mean of KUALITAS': [int(mean_selected_columns['KUALITAS'])],
            'Mean of K3': [int(mean_selected_columns['K3'])],
            'Mean of L': [int(mean_selected_columns['L'])],
            'Mean of TOT POINT': [int(mean_selected_columns['TOT POINT'])]
        })
        def assign_grade(score):
            if score <= 40:
                return 'E'
            elif score <= 60:
                return 'D'
            elif score <= 70:
                return 'C'
            elif score <= 89:
                return 'B'
            else:
                return 'A'

        summary_df['Nilai'] = summary_df['Mean of TOT POINT'].apply(assign_grade)
        
        def assign_grade(score):
            if score <= 40:
                return 'Sangat Buruk / tidak dilanjutkan kerjasamanya'
            elif score <= 60:
                return 'Buruk / dipertimbangkan kerjasamanya dan diberikan surat peringatan'
            elif score <= 70:
                return 'Cukup / dilanjutkan kerjasamanya'
            elif score <= 89:
                return 'Baik / dilanjutkan kerjasamanya'
            else:
                return 'Sangat Baik / dilanjutkan kerjasamanya'

        summary_df['Keterangan'] = summary_df['Mean of TOT POINT'].apply(assign_grade)

        st.write("Summary:")
        st.dataframe(summary_df)
                
        doc = Document("templates/Template.docx")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Ganti teks di dalam sel tabel
                    if "(Vendor_name)" in cell.text:
                        cell.text = cell.text.replace("(Vendor_name)", filtered_data['Vendor_name'].iloc[0])
                    elif '(PP)' in cell.text:
                        cell.text = cell.text.replace('(PP)', str(int(summary_df['Mean of RP/PP'].iloc[0])))
                    elif '(K)' in cell.text:
                        cell.text = cell.text.replace('(K)', str(int(summary_df['Mean of KUALITAS'].iloc[0])))
                    elif '(K3)' in cell.text:
                        cell.text = cell.text.replace('(K3)', str(int(summary_df['Mean of K3'].iloc[0])))
                    elif '(L)' in cell.text:
                        cell.text = cell.text.replace('(L)', str(int(summary_df['Mean of L'].iloc[0])))
        
        for paragraph in doc.paragraphs:
            if "{Keterangan}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{Keterangan}", summary_df['Keterangan'].iloc[0])
            if "(TOT)" in paragraph.text:
                paragraph.text = paragraph.text.replace("(TOT)", str(int(summary_df['Mean of TOT POINT'].iloc[0])))
        
        
# Load data from the uploaded file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read all sheets in the Excel file and concatenate them into one DataFrame
    result_df = read_all_sheets_in_excel(uploaded_file)

    # Display filtered data based on user selection
    display_filtered_data(result_df)

    if st.button("Download Word Document"):
        if doc is not None:
            byte_io = BytesIO()
            doc.save(byte_io)
            byte_io.seek(0)
            st.download_button(label="Download Here", data=byte_io, file_name="Result.docx")

