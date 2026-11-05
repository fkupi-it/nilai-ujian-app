import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import random
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Sistem Pengolahan Ujian FK UPI", layout="wide")

# Helper: Format Tanggal ke Bahasa Indonesia (Contoh: 30 Oktober 2026)
NAMA_BULAN = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

# Mapping Jenis Ujian -> Kode Ujian
MAP_KODE_UJIAN = {
    "Post Test": "PO",
    "Pre Test": "PR",
    "Remedial Post Test": "POR",
    "UTS": "UT",
    "UAS": "UA",
    "Remedial": "RE"
}

def format_tanggal_indonesia(dt_date):
    day = dt_date.day
    month = NAMA_BULAN[dt_date.month]
    year = dt_date.year
    return f"{day} {month} {year}"

# 1. Load Master Data Mahasiswa (Global)
def load_master_data():
    try:
        df = pd.read_excel("master_mahasiswa.xlsx")
        df.columns = [str(c).strip() for c in df.columns]
        
        rename_master = {}
        for c in df.columns:
            c_clean = c.lower().strip()
            if c_clean in ['nim', 'nomor peserta', 'no peserta']:
                rename_master[c] = 'NIM'
            elif c_clean in ['nama', 'nama mahasiswa', 'nama peserta']:
                rename_master[c] = 'Nama'
            elif c_clean in ['angkatan']:
                rename_master[c] = 'Angkatan'
                
        df = df.rename(columns=rename_master)
        
        if 'NIM' in df.columns:
            df['NIM'] = df['NIM'].astype(str).str.strip()
        if 'Angkatan' in df.columns:
            df['Angkatan'] = df['Angkatan'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
        return df
    except FileNotFoundError:
        st.error("❌ File 'master_mahasiswa.xlsx' tidak ditemukan di repository GitHub!")
        return None
    except Exception as e:
        st.error(f"❌ Gagal membaca 'master_mahasiswa.xlsx'. Detail Error: {e}")
        return None

df_master_all = load_master_data()

# --- SYSTEM NAVIGATION MENU ---
st.sidebar.title("📌 Menu Navigasi")
menu_option = st.sidebar.radio("Pilih Layanan:", ["📊 Olah Nilai", "⚡ Generate Soal"])

# ==============================================================================
# MENU 1: OLAH NILAI
# ==============================================================================
if menu_option == "📊 Olah Nilai":
    st.title("📊 Pemroses Nilai & Analisis Ujian FK UPI")
    
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
            default_index = list_angkatan.index("2023") if "2023" in list_angkatan else 0
            selected_angkatan = st.selectbox("Pilih Angkatan", list_angkatan, index=default_index)
        else:
            selected_angkatan = st.text_input("Angkatan", value="2023").strip()

    st.divider()

    st.subheader("2. Unggah File Ujian")
    file_ujian = st.file_uploader("Upload File Hasil Ujian (.xlsx)", type=["xlsx"], key="olah_nilai")

    if file_ujian and df_master_all is not None:
        if 'Angkatan' in df_master_all.columns:
            df_master = df_master_all[df_master_all['Angkatan'] == str(selected_angkatan)].copy()
        else:
            df_master = df_master_all.copy()

        if df_master.empty:
            st.warning(f"⚠️ Tidak ada data mahasiswa angkatan {selected_angkatan} di dalam file master_mahasiswa.xlsx!")
            st.stop()

        try:
            df_raw = pd.read_excel(file_ujian)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            
            rename_dict = {}
            for col in df_raw.columns:
                col_clean = col.lower().replace('_', ' ').replace('-', ' ').strip()
                if col_clean in ['nomor peserta', 'no peserta', 'nim', 'no. peserta']:
                    rename_dict[col] = 'NIM'
                elif col_clean in ['nama peserta', 'nama mahasiswa', 'nama', 'full name']:
                    rename_dict[col] = 'Nama'
            
            df_raw = df_raw.rename(columns=rename_dict)
            
            if 'NIM' not in df_raw.columns:
                st.error("❌ Kolom 'Nomor Peserta' atau 'NIM' tidak ditemukan di file ujian.")
                st.stop()
                
            df_raw['NIM'] = df_raw['NIM'].astype(str).str.strip()
            
            if 'Nama' not in df_raw.columns and 'Nama' in df_master.columns:
                map_nama = dict(zip(df_master['NIM'], df_master['Nama']))
                df_raw['Nama'] = df_raw['NIM'].map(map_nama).fillna('')

        except Exception as e:
            st.error(f"❌ Gagal membaca file ujian yang diunggah. Detail: {e}")
            st.stop()
        
        non_soal_cols = {'NIM', 'Nama', 'Sesi', 'Nomor Peserta', 'Nama Peserta'}
        soal_cols = [c for c in df_raw.columns if c not in non_soal_cols]
        
        if 'Sesi' not in df_raw.columns:
            df_raw['Sesi'] = "1"
        
        kunci = 'A'
        
        df_raw['Jumlah Benar'] = (df_raw[soal_cols] == kunci).sum(axis=1)
        total_soal = len(soal_cols)
        df_raw['Nilai (100)'] = ((df_raw['Jumlah Benar'] / total_soal) * 100).round(2)
        
        nim_hadir = set(df_raw['NIM'])
        df_th = df_master[~df_master['NIM'].isin(nim_hadir)].copy()
        
        cols_nilai_order = ['Sesi', 'NIM', 'Nama', 'Jumlah Benar', 'Nilai (100)'] + soal_cols
        
        if not df_th.empty:
            df_th['Sesi'] = "TH"
            df_th['Jumlah Benar'] = 0
            df_th['Nilai (100)'] = 0.0
            for sc in soal_cols:
                df_th[sc] = "TH"
                
            df_nilai_th = df_th.reindex(columns=cols_nilai_order).fillna('')
        else:
            df_nilai_th = pd.DataFrame(columns=cols_nilai_order)
                
        df_nilai_hadir = df_raw.reindex(columns=cols_nilai_order).fillna('')
        
        if not df_th.empty:
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
        
        # Output Excel
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df_nilai_final.to_excel(writer, sheet_name='Nilai', index=False)
            
            df_metadata = pd.DataFrame([
                ["Nama Mata Kuliah:", nama_matkul],
                ["Jenis Ujian:", jenis_ujian],
                ["Versi Soal:", versi_soal],
                ["Angkatan:", str(selected_angkatan)],
                ["Tanggal Ujian:", str(tgl_ujian)],
                ["Jumlah Soal:", total_soal],
                ["Peserta Hadir / TH:", f"{peserta_hadir} / {peserta_th}"]
            ])
            df_metadata.to_excel(writer, sheet_name='Analysis', index=False, header=False, startrow=0)
            df_analysis_table.to_excel(writer, sheet_name='Analysis', index=False, startrow=8)
            
            wb = writer.book
            
            ws_nilai = wb['Nilai']
            ws_nilai.column_dimensions['A'].width = 15
            ws_nilai.column_dimensions['B'].width = 15
            ws_nilai.column_dimensions['C'].width = 40
            ws_nilai.column_dimensions['D'].width = 16
            ws_nilai.column_dimensions['E'].width = 16
            
            for r in range(2, len(df_nilai_final) + 2):
                ws_nilai.cell(row=r, column=5).number_format = '0.00'
                
            for col_idx in range(6, 6 + len(soal_cols)):
                col_letter = get_column_letter(col_idx)
                ws_nilai.column_dimensions[col_letter].width = 8
            
            ws_analysis = wb['Analysis']
            col_widths_analysis = {
                'A': 22, 'B': 15, 'C': 22, 'D': 16, 'E': 22, 'F': 20, 'G': 26,
                'H': 8, 'I': 8, 'J': 8, 'K': 8, 'L': 8
            }
            for col_letter, width in col_widths_analysis.items():
                ws_analysis.column_dimensions[col_letter].width = width
                
            fill_purple = PatternFill(start_color="A7A6F6", end_color="A7A6F6", fill_type="solid")
            fill_blue   = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            fill_green  = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            fill_yellow = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
            fill_pink   = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
            
            start_row_data = 10
            end_row_data = start_row_data + len(df_analysis_table) - 1
            
            for r in range(start_row_data, end_row_data + 1):
                for col_idx in [2, 3, 4, 6, 8, 9, 10, 11, 12]:
                    cell = ws_analysis.cell(row=r, column=col_idx)
                    cell.number_format = '0.00'
                
                cell_diff = ws_analysis.cell(row=r, column=5)
                val_diff = str(cell_diff.value)
                if val_diff == "Very Easy": cell_diff.fill = fill_purple
                elif val_diff == "Easy": cell_diff.fill = fill_blue
                elif val_diff == "Moderate": cell_diff.fill = fill_green
                elif val_diff == "Difficult": cell_diff.fill = fill_yellow
                elif val_diff == "Very Difficult": cell_diff.fill = fill_pink
                
                cell_disc = ws_analysis.cell(row=r, column=7)
                val_disc = str(cell_disc.value)
                if val_disc == "Very good item": cell_disc.fill = fill_green
                elif val_disc == "To be revised": cell_disc.fill = fill_yellow
                elif val_disc == "To be discarded": cell_disc.fill = fill_pink
                
                for c_opt in range(8, 13):
                    cell_opt = ws_analysis.cell(row=r, column=c_opt)
                    try:
                        val_opt = float(cell_opt.value)
                        if val_opt < 0.05:
                            cell_opt.fill = fill_pink
                    except (ValueError, TypeError):
                        pass

        output_buffer.seek(0)
        
        tgl_yymmdd = tgl_ujian.strftime("%y%m%d")
        combo_tgl_versi = f"{tgl_yymmdd}{versi_soal}"
        tgl_indo = format_tanggal_indonesia(tgl_ujian)
        
        filename_download = f"[Nilai] {nama_matkul}_{combo_tgl_versi}_{jenis_ujian}_{tgl_indo}.xlsx"
        
        st.success(f"✅ Data berhasil diproses untuk **{nama_matkul} - {jenis_ujian}**!")
        st.download_button(
            label=f"📥 Download File: {filename_download}",
            data=output_buffer,
            file_name=filename_download,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==============================================================================
# MENU 2: GENERATE SOAL (PENGISIAN BERBASIS TEMPLATE ASLI EXAMTOOLS)
# ==============================================================================
elif menu_option == "⚡ Generate Soal":
    st.title("⚡ Generate Import Soal CBT")
    
    # 1. INFORMASI UJIAN
    st.subheader("1. Informasi Ujian")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        matkul_gen = st.text_input("Mata Kuliah (Kode/Singkatan)", value="FBS1").strip().upper()
        jenis_gen = st.selectbox(
            "Jenis Ujian", 
            ["Post Test", "Pre Test", "Remedial Post Test", "UTS", "UAS", "Remedial"], 
            index=0
        )
        kode_gen = MAP_KODE_UJIAN.get(jenis_gen, "PO")
        st.caption(f"Kode Ujian Otomatis: **{kode_gen}**")

    with col2:
        tgl_gen_input = st.date_input("Tanggal Ujian", value=datetime.date.today())
        tgl_gen_yymmdd = tgl_gen_input.strftime("%y%m%d")
        st.caption(f"Format ID Tanggal: **{tgl_gen_yymmdd}**")

    with col3:
        versi_gen = st.selectbox("Versi Soal", ["1", "2", "3", "4", "5"], index=0)

    st.divider()

    # 2. UNGGAH FILE SOAL
    st.subheader("2. Unggah File Soal")
    st.caption("Unggah file Excel (hanya 1 sheet) dengan urutan kolom: **1. No.Soal | 2. Vignette | 3. Pertanyaan | 4. Media | 5. Opsi Benar | 6. Opsi Salah 1 | 7. Opsi Salah 2 | 8. Opsi Salah 3 | 9. Opsi Salah 4**")
    
    file_soal = st.file_uploader("Upload File Excel Soal (.xlsx)", type=["xlsx"], key="gen_soal")
    
    if file_soal:
        try:
            # Buka Template Master ExamTools
            try:
                wb_tpl = openpyxl.load_workbook("template_examtools.xlsx")
                ws_tpl = wb_tpl.active
            except FileNotFoundError:
                st.error("❌ File 'template_examtools.xlsx' tidak ditemukan di repositori GitHub! Mohon simpan file template di folder repositori.")
                st.stop()

            df_soal = pd.read_excel(file_soal, sheet_name=0)
            
            if df_soal.empty:
                st.warning("⚠️ File Excel soal tidak berisi data.")
                st.stop()
                
            current_row = 2 # Mulai mengisi dari baris ke-2 (di bawah header)
            count_soal = 0
            
            for idx, row in df_soal.iterrows():
                def get_by_idx(col_idx):
                    if col_idx < len(row):
                        val = row.iloc[col_idx]
                        if pd.notna(val):
                            return str(val).strip()
                    return ""

                no_soal_raw = get_by_idx(0)
                if not no_soal_raw:
                    no_soal_raw = str(idx + 1)
                if no_soal_raw.endswith('.0'):
                    no_soal_raw = no_soal_raw[:-2]
                    
                raw_vignette = get_by_idx(1)
                pertanyaan   = get_by_idx(2)
                media        = get_by_idx(3).lower()
                
                opsi_benar  = get_by_idx(4)
                opsi_salah1 = get_by_idx(5)
                opsi_salah2 = get_by_idx(6)
                opsi_salah3 = get_by_idx(7)
                opsi_salah4 = get_by_idx(8)
                
                if not no_soal_raw and not raw_vignette and not pertanyaan:
                    continue
                    
                no_soal_formatted = no_soal_raw.zfill(3)
                
                id_base = f"{matkul_gen}-{kode_gen}-{tgl_gen_yymmdd}{versi_gen}{no_soal_formatted}"
                random_two_digits = str(random.randint(10, 99))
                id_soal_lengkap = f"{id_base}{random_two_digits}"
                
                vignette_formatted = f'<span style="font-size: 0.8em; color: #ebebeb;">{id_soal_lengkap}</span><br>{raw_vignette}'
                
                if 'gambar' in media:
                    vignette_formatted += f'<br><br><img src="https://fkupi-it.github.io/mediacbt/{id_base}.jpg" alt="Gambar soal" style="max-width:100%; height:auto;">'
                if 'audio' in media:
                    vignette_formatted += f'<br><br><video controls="controls" width="300" height="150"><source src="https://fkupi-it.github.io/mediacbt/{id_base}.mp3" type="audio/mpeg"></video>'
                
                # Pengisian Sel secara langsung ke Template openpyxl (Presisi Float & Null)
                ws_tpl.cell(row=current_row, column=1, value=0.0)  # Tipe Soal
                ws_tpl.cell(row=current_row, column=2, value=1.0)  # Tampilan Urutan Pilihan
                ws_tpl.cell(row=current_row, column=3, value=0.0)  # Sembunyikan Pilihan
                ws_tpl.cell(row=current_row, column=5, value=0.0)  # Enkripsi
                
                ws_tpl.cell(row=current_row, column=16, value=vignette_formatted) # Badan Soal
                ws_tpl.cell(row=current_row, column=17, value=pertanyaan)          # Pertanyaan
                ws_tpl.cell(row=current_row, column=18, value=1.0)                 # Kunci
                
                if opsi_benar:  ws_tpl.cell(row=current_row, column=19, value=opsi_benar)
                if opsi_salah1: ws_tpl.cell(row=current_row, column=20, value=opsi_salah1)
                if opsi_salah2: ws_tpl.cell(row=current_row, column=21, value=opsi_salah2)
                if opsi_salah3: ws_tpl.cell(row=current_row, column=22, value=opsi_salah3)
                if opsi_salah4: ws_tpl.cell(row=current_row, column=23, value=opsi_salah4)
                
                current_row += 1
                count_soal += 1
                
            gen_buffer = io.BytesIO()
            wb_tpl.save(gen_buffer)
            gen_buffer.seek(0)
            
            combo_tgl_versi_gen = f"{tgl_gen_yymmdd}{versi_gen}"
            filename_gen = f"[ExamTools - Soal] {matkul_gen}_{combo_tgl_versi_gen}_{jenis_gen}.xlsx"
            
            st.success(f"✅ Berhasil mengonversi **{count_soal} soal** ke dalam template ExamTools!")
            st.download_button(
                label=f"📥 Download File Import CBT: {filename_gen}",
                data=gen_buffer,
                file_name=filename_gen,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses file soal. Detail: {e}")