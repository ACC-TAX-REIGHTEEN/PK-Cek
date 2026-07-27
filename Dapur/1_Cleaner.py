import os
import re
import pandas as pd

def normalize_text(text):
    if pd.isna(text):
        return ""
    clean_str = str(text).replace('\xa0', ' ').strip()
    clean_str = re.sub(r'\s+', ' ', clean_str)
    return clean_str

def get_standard_col_name(combined_text):
    norm = normalize_text(combined_text).lower()
    
    if "tanggal" in norm and "pajak" not in norm:
        return "Tanggal"
    elif "tgl" in norm and "pajak" in norm:
        return "Tgl. Pajak"
    elif "referensi" in norm or "ref" in norm:
        return "No. Referensi"
    elif "faktur" in norm:
        return "No. Faktur Pajak"
    elif "pelanggan" in norm or "cust" in norm:
        if "nama" in norm:
            return "Nama Pelanggan"
        elif "negara" in norm:
            return "Negara Pelanggan"
        elif "pajak" in norm or "npwp" in norm:
            return "Nomor Pajak Pelanggan"
        else:
            return "No. Pelanggan"
    elif "pajak" in norm and "jumlah" in norm:
        return "Jumlah Pajak"
    
    return None

def clean_excel_pk(input_filepath="PK.xls", output_filepath="PK_temp.xlsx"):
    if not os.path.exists(input_filepath):
        if os.path.exists("PK.xlsx"):
            input_filepath = "PK.xlsx"
        else:
            raise FileNotFoundError(f"File '{input_filepath}' tidak ditemukan.")

    print(f"Membaca file: {input_filepath}...")
    
    df_raw = pd.read_excel(input_filepath, header=None)
    
    col_mapping = {}
    top_rows = df_raw.iloc[:12]
    
    for c_idx in range(df_raw.shape[1]):
        col_cells = top_rows.iloc[:, c_idx].dropna().astype(str)
        combined_header_text = " ".join(col_cells)
        
        std_name = get_standard_col_name(combined_header_text)
        if std_name and std_name not in col_mapping.values():
            col_mapping[c_idx] = std_name

    print("Kolom terdeteksi:", list(col_mapping.values()))
    
    if not col_mapping:
        raise ValueError("Header kolom tidak berhasil terdeteksi. Periksa kembali file Excel.")

    selected_indices = list(col_mapping.keys())
    selected_names = list(col_mapping.values())
    
    df_data = df_raw.iloc[:, selected_indices].copy()
    df_data.columns = selected_names
    
    df_cleaned = df_data.dropna(how='all').copy()
    
    ref_col = "Nama Pelanggan" if "Nama Pelanggan" in df_cleaned.columns else df_cleaned.columns[0]
    
    df_cleaned = df_cleaned[
        df_cleaned[ref_col].notna() & 
        (df_cleaned[ref_col].astype(str).str.strip() != "") &
        (~df_cleaned[ref_col].astype(str).str.contains("Nama Pelanggan|Pelanggan|Saldo Awal|Tanggal", case=False, na=False))
    ]
    
    def clean_code_format(val):
        if pd.isna(val):
            return ""
        s = normalize_text(val)
        s = re.sub(r'[,.]00$', '', s)
        return s

    def clean_no_pelanggan(val):
        if pd.isna(val):
            return ""
        s = normalize_text(val)
        s = re.sub(r'[,.]00$', '', s)
        s = s.replace('.', '').replace(',', '')
        return s

    def clean_jumlah_pajak(val):
        if pd.isna(val):
            return ""
        s = normalize_text(val)
        s = re.sub(r'[,.]00$', '', s)
        return s

    if "No. Referensi" in df_cleaned.columns:
        df_cleaned["No. Referensi"] = df_cleaned["No. Referensi"].apply(clean_code_format)
        
    if "No. Faktur Pajak" in df_cleaned.columns:
        df_cleaned["No. Faktur Pajak"] = df_cleaned["No. Faktur Pajak"].apply(clean_code_format)
        
    if "Nomor Pajak Pelanggan" in df_cleaned.columns:
        df_cleaned["Nomor Pajak Pelanggan"] = df_cleaned["Nomor Pajak Pelanggan"].apply(clean_code_format)
        
    if "No. Pelanggan" in df_cleaned.columns:
        df_cleaned["No. Pelanggan"] = df_cleaned["No. Pelanggan"].apply(clean_no_pelanggan)
        
    if "Jumlah Pajak" in df_cleaned.columns:
        df_cleaned["Jumlah Pajak"] = df_cleaned["Jumlah Pajak"].apply(clean_jumlah_pajak)

    df_cleaned.to_excel(output_filepath, index=False)
    print(f"Pembersihan selesai! Data berhasil disimpan di: '{output_filepath}'")
    print(f"Total baris data transaksi: {len(df_cleaned)} baris.")
    
    return df_cleaned

if __name__ == "__main__":
    clean_excel_pk()