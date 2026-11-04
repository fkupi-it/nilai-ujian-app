import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime

st.set_page_config(page_title="Pemroses Nilai & Analisis Ujian", layout="wide")

st.title("📊 Pemroses Nilai & Analisis Ujian FK UPI")

# 1. Load Master Data Mahasiswa (Tanpa Cache agar tidak stuck)
def load_master_data():
    try:
        df = pd.read_excel("master_mahasiswa.xlsx")
        # Rapikan nama kolom (hilangkan spasi tak sengaja)
        df.columns = [str(c).strip() for c in df.columns]
        
        if 'NIM' in df.columns:
            df['NIM'] = df['NIM'].astype(str).str.strip()
        if 'Angkatan' in df.columns:
            df['Angkatan'] = df['Angkatan'].astype(str).str.strip()
            
        return df
    except FileNotFoundError:
        st.error("❌ File 'master_mahasiswa.xlsx' tidak ditemukan di repository GitHub! Pastikan file berada di folder utama (root).")
        return None
    except Exception as e:
        st.error(f"❌ Gagal membaca 'master_mahasiswa.xlsx'. Detail Error: {e}")
        return None

df_master_all = load_master_data()

# 2. Form Metadata Ujian
st.subheader("1. Informasi Ujian")
col1, col2, col3 = st.columns(3)

with col1:
    nama_matkul = st.text_input("Nama Matkul (Kode/Singkatan)", value="FBS1").strip().upper()
    jenis_ujian = st.text_input("Jenis Ujian", value="Post Test").strip()

with col2:
    versi_soal = st.text_input("Versi Soal", value="1").strip()
    tgl_ujian = st.date_input("Tanggal Ujian", value=datetime.date.today())

with col3:
    if df_master_all is not None and 'Angkatan' in df_master_all.columns:
        list_angkatan = sorted(list(df_master_all['Angkatan'].dropna().unique()))
        # Tentukan index default untuk 2023 jika ada dalam daftar
        default_index = list_angkatan.index("2023") if "2023" in list_angkatan else 0
        selected_angkatan = st.selectbox("Pilih Angkatan", list_angkatan, index=default_index)
    else:
        selected_angkatan = st.text_input("Angkatan", value="2023").strip()

st.divider()

# 3. Form Upload File
st.subheader("2. Unggah File Ujian")
file_ujian = st.file_uploader("Upload File Hasil Ujian (.xlsx)", type=["xlsx"])

# Jalankan kalkulasi hanya jika file master ada & file ujian sudah di-upload
if file_ujian and df_master_all is not None:
    # Filter Master Berdasarkan Angkatan
    if 'Angkatan' in df_master_all.columns:
        df_master = df_master_all[df_master_all['Angkatan'] == str(selected_angkatan)].copy()
    else:
        df_master = df_master_all.copy()

    if df_master.empty:
        st.warning(f"⚠️ Tidak ada data mahasiswa angkatan {selected_angkatan} di dalam file master_mahasiswa.xlsx!")
        st.stop()

    # Read Excel Ujian
    try:
        df_raw = pd.read_excel(file_ujian)
        df_raw['NIM'] = df_raw['NIM'].astype(str).str.strip()
    except Exception as e:
        st.error(f"❌ Gagal membaca file ujian yang diunggah. Detail: {e}")
        st.stop()
    
    # Deteksi Kolom Soal (Q1 - Qn)
    cols_base = ['NIM', 'Nama', 'Sesi']
    soal_cols = [c for c in df_raw.columns if c not in cols_base]
    
    kunci = 'A'
    
    # Hitung Nilai
    df_raw['Jumlah Benar'] = (df_raw[soal_cols] == kunci).sum(axis=1)
    total_soal = len(soal_cols)
    df_raw['Nilai (100)'] = ((df_raw['Jumlah Benar'] / total_soal) * 100).round(2)
    
    # Matching Tracing TH
    nim_hadir = set(df_raw['NIM'])
    df_th = df_master[~df_master['NIM'].isin(nim_hadir)].copy()
    
    if not df_th.empty:
        df_th['Sesi'] = "TH"
        df_th['Jumlah Benar'] = 0
        df_th['Nilai (100)'] = 0.0
        for sc in soal_cols:
            df_th[sc] = "TH"
            
    cols_nilai_order = ['Sesi', 'NIM', 'Nama', 'Jumlah Benar', 'Nilai (100)'] + soal_cols
    df_nilai_hadir = df_raw[cols_nilai_order]
    
    if not df_th.empty:
        df_nilai_th = df_th[cols_nilai_order]
        df_nilai_final = pd.concat([df_nilai_hadir, df_nilai_th], ignore_index=True)
    else:
        df_nilai_final = df_nilai_hadir.copy()
        
    df_nilai_final = df_nilai_final.sort_values(by=['Nilai (100)', 'NIM'], ascending=[False, True]).reset_index(drop=True)
    
    # Analisis Statistik
    opsi_valid = ['A', 'B', 'C', 'D', 'E']
    df_analisis_peserta = df_raw[df_raw[soal_cols].isin(opsi_valid).any(axis=1)].copy()
    
    total_terdaftar = len(df_master)
    peserta_hadir = len(df_analisis_peserta)
    peserta_th = total_terdaftar - peserta_hadir
    
    df_analisis_peserta = df_analisis_peserta.sort_values(by='Nilai (100)', ascending=False).reset_index(drop=True)
    n_group = int(np.ceil(peserta_hadir * 0.27)) if peserta_hadir > 0 else 0
    top_group = df_analisis_peserta.head(n_group)
    bottom_group = df_analisis_peserta.tail(n_group)
    T = n_group * 2
    
    analysis_rows = []
    for soal in soal_cols:
        jumlah_benar = (df_analisis_peserta[soal] == 'A').sum()
        persentase = (jumlah_benar / peserta_hadir * 100) if peserta_hadir > 0 else 0
        
        Ru = (top_group[soal] == 'A').sum() if n_group > 0 else 0
        Rl = (bottom_group[soal] == 'A').sum() if n_group > 0 else 0
        
        p = (Ru + Rl) / T if T > 0 else 0
        D = (Ru - Rl) / (0.5 * T) if T > 0 else 0
        
        if p > 0.85: interp_diff = "Very Easy"
        elif p >= 0.71: interp_diff = "Easy"
        elif p >= 0.30: interp_diff = "Moderate"
        elif p >= 0.15: interp_diff = "Difficult"
        else: interp_diff = "Very Difficult"
        
        if D > 0.85: interp_disc = "To be discarded"
        elif D >= 0.71: interp_disc = "To be revised"
        elif D >= 0.30: interp_disc = "Very good item"
        elif D >= 0.15: interp_disc = "To be revised"
        else: interp_disc = "To be discarded"
        
        proporsi_opsi = []
        for opt in ['A', 'B', 'C', 'D', 'E']:
            c_top = (top_group[soal] == opt).sum() if n_group > 0 else 0
            c_bot = (bottom_group[soal] == opt).sum() if n_group > 0 else 0
            tot_opt = (c_top + c_bot) / T if T > 0 else 0
            proporsi_opsi.append(round(tot_opt, 2))
            
        row = [soal, jumlah_benar, round(persentase, 2), round(p, 2), interp_diff, round(D, 2), interp_disc] + proporsi_opsi
        analysis_rows.append(row)
        
    analysis_headers = [
        "Soal", "Jumlah Benar", "Persentase Benar (%)", "Difficulty Index",
        "Interpretasi Difficulty", "Discrimination Index", "Interpretasi Discrimination",
        "A", "B", "C", "D", "E"
    ]
    df_analysis_table = pd.DataFrame(analysis_rows, columns=analysis_headers)
    
    # Generate Output Excel
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df_nilai_final.to_excel(writer, sheet_name='Nilai', index=False)
        
        df_metadata = pd.DataFrame([
            ["Nama Mata Kuliah:", nama_matkul],
            ["Jenis Ujian:", jenis_ujian],
            ["Versi Soal:", versi_soal],
            ["Angkatan:", selected_angkatan],
            ["Tanggal Ujian:", str(tgl_ujian)],
            ["Jumlah Soal:", total_soal],
            ["Peserta Hadir / TH:", f"{peserta_hadir} / {peserta_th}"]
        ])
        df_metadata.to_excel(writer, sheet_name='Analysis', index=False, header=False, startrow=0)
        df_analysis_table.to_excel(writer, sheet_name='Analysis', index=False, startrow=8)
        
    output_buffer.seek(0)
    
    filename_download = f"Hasil_{nama_matkul}_{jenis_ujian}_{selected_angkatan}_V{versi_soal}_{tgl_ujian}.xlsx"
    
    st.success(f"✅ Data berhasil diproses untuk **{nama_matkul} - {jenis_ujian} (Angkatan {selected_angkatan})**!")
    st.download_button(
        label=f"📥 Download File: {filename_download}",
        data=output_buffer,
        file_name=filename_download,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )