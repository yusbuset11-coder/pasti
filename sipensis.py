import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io

# Konfigurasi Google Sheets API
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client():
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    return None

def run_sipensis():
    st.markdown("## 📊 SIPENSIS: Sistem Informasi Presensi Siswa")
    
    # Identifikasi User yang Sedang Login
    user_email = st.session_state.get("user_email", "budi@gmail.com")
    user_name = st.session_state.get("user_name", "Budi Santoso")
    
    client = get_gspread_client()
    if not client:
        st.error("Koneksi Google Sheets (GCP Service Account) belum dikonfigurasi di secrets.toml.")
        return
    
    # Hubungkan ke DATA_MASTER_REGISTRY
    MASTER_REGISTRY_ID = st.secrets.get("master_registry_id", "1mgN63xzrLt__5b9-gBw8dIWYP3RRgNdagUiTurFZdgg")
    
    try:
        registry_sheet = client.open_by_key(MASTER_REGISTRY_ID).sheet1
        registry_data = registry_sheet.get_all_records()
        df_registry = pd.DataFrame(registry_data)
    except Exception as e:
        st.error(f"Gagal memuat DATA_MASTER_REGISTRY: {e}")
        return
    
    # Cari data guru berdasarkan email
    guru_row_idx = None
    guru_spreadsheet_id = None
    
    for idx, row in df_registry.iterrows():
        if str(row.get("Email", "")).strip().lower() == str(user_email).strip().lower():
            guru_row_idx = idx + 2  # Baris header sheet
            guru_spreadsheet_id = str(row.get("Spreadsheet_ID_Guru", "")).strip()
            break
            
    if guru_row_idx is None:
        st.warning(f"Akun dengan email {user_email} belum terdaftar di DATA_MASTER_REGISTRY.")
        return
        
    # Jika Spreadsheet_ID_Guru kosong
    if not guru_spreadsheet_id or guru_spreadsheet_id in ["(Kosongkan dulu)", "nan", ""]:
        st.info("💡 **Setup Awal Database Pribadi Guru**\nSilakan buat salinan dari Template_Database_Guru, lalu masukkan Spreadsheet ID milik Anda di bawah ini agar data absensi tersimpan di akun Anda sendiri.")
        new_sheet_id = st.text_input("Masukkan Spreadsheet ID Pribadi Anda:")
        if st.button("Simpan Spreadsheet ID"):
            if new_sheet_id:
                registry_sheet.update_cell(guru_row_idx, 5, new_sheet_id)  # Kolom E
                st.success("Spreadsheet ID berhasil disimpan! Silakan refresh halaman.")
                st.rerun()
            else:
                st.warning("Mohon masukkan ID Spreadsheet yang valid.")
        return

    # Koneksi ke Spreadsheet Pribadi Guru
    try:
        teacher_wb = client.open_by_key(guru_spreadsheet_id)
        sheet_absensi = teacher_wb.worksheet("Absensi Harian")
        sheet_rekap_ganjil = teacher_wb.worksheet("Rekap Semester Ganjil")
        sheet_rekap_genap = teacher_wb.worksheet("Rekap Semester Genap")
    except Exception as e:
        st.error(f"Gagal mengakses Spreadsheet Pribadi Guru. Pastikan Service Account sudah diberi akses Editor. Detail: {e}")
        return

    # Ambil data Absensi Harian untuk filter
    data_absensi = sheet_absensi.get_all_records()
    df_absensi = pd.DataFrame(data_absensi)
    
    # Filter Sekolah & Kelas di Sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🏫 Filter Wilayah & Kelas")
    
    list_sekolah = ["Semua Sekolah"]
    list_kelas = ["Semua Kelas"]
    
    if not df_absensi.empty and "Sekolah" in df_absensi.columns and "Kelas" in df_absensi.columns:
        list_sekolah += list(df_absensi["Sekolah"].dropna().unique())
        list_kelas += list(df_absensi["Kelas"].dropna().unique())
        
    pilih_sekolah = st.sidebar.selectbox("Pilih Sekolah", list_sekolah)
    pilih_kelas = st.sidebar.selectbox("Pilih Kelas", list_kelas)

    # Menu Navigasi Tabs SIPENSIS
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📝 Input Manual", 
        "📥 Download & Upload Database Guru", 
        "📈 Laporan Harian", 
        "📊 Rekap Semester Ganjil", 
        "📊 Rekap Semester Genap"
    ])
    
    with tab1:
        st.subheader("Form Input Presensi Harian Siswa")
        
        if df_absensi.empty:
            st.warning("Belum ada data siswa di database Anda. Silakan upload template melalui tab 'Download & Upload Database Guru' terlebih dahulu.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                tanggal_absen = st.date_input("Tanggal Absensi", datetime.now())
                nama_guru_input = st.text_input("Nama Guru", value=user_name, disabled=True)
            with col2:
                # 1. Pilihan Sekolah Interaktif (Dropdown)
                list_sekolah_db = df_absensi["Sekolah"].dropna().unique().tolist() if "Sekolah" in df_absensi.columns else []
                pilih_sekolah_form = st.selectbox("Pilih Sekolah", list_sekolah_db if list_sekolah_db else ["-"])
                
                # Filter Kelas berdasarkan Sekolah yang dipilih
                df_filter_sekolah = df_absensi[df_absensi["Sekolah"] == pilih_sekolah_form] if list_sekolah_db else pd.DataFrame()
                list_kelas_db = df_filter_sekolah["Kelas"].dropna().unique().tolist() if "Kelas" in df_filter_sekolah.columns else []
                pilih_kelas_form = st.selectbox("Pilih Kelas", list_kelas_db if list_kelas_db else ["-"])
                
            # Input mata pelajaran secara dinamis
            mata_pelajaran = st.text_input("Mata Pelajaran yang Diajar", placeholder="Contoh: Pendidikan Pancasila")
            
            st.markdown("---")
            st.markdown(f"### Daftar Siswa Kelas {pilih_kelas_form} ({pilih_sekolah_form})")
            
            # Saring daftar siswa unik berdasarkan sekolah dan kelas yang dipilih
            df_filter_kelas = df_filter_sekolah[df_filter_sekolah["Kelas"] == pilih_kelas_form] if not df_filter_sekolah.empty else pd.DataFrame()
            df_siswa_unik = df_filter_kelas[["No", "Nama_Siswa"]].drop_duplicates().sort_values(by="No") if not df_filter_kelas.empty and "No" in df_filter_kelas.columns and "Nama_Siswa" in df_filter_kelas.columns else pd.DataFrame()
            
            absensi_rows = []
            if not df_siswa_unik.empty:
                for idx, row_s in df_siswa_unik.iterrows():
                    no_s = row_s.get("No", idx + 1)
                    nama_s = row_s.get("Nama_Siswa", "")
                    
                    cols = st.columns([1, 4, 1, 1, 1])
                    with cols[0]:
                        st.write(f"No. {no_s}")
                    with cols[1]:
                        st.write(nama_s)
                    with cols[2]:
                        s_checked = st.checkbox("S", key=f"s_{pilih_sekolah_form}_{pilih_kelas_form}_{no_s}")
                    with cols[3]:
                        i_checked = st.checkbox("I", key=f"i_{pilih_sekolah_form}_{pilih_kelas_form}_{no_s}")
                    with cols[4]:
                        a_checked = st.checkbox("A", key=f"a_{pilih_sekolah_form}_{pilih_kelas_form}_{no_s}")
                    
                    # 2. Simpan sebagai string "V" atau kosong "" (Bukan True/False)
                    absensi_rows.append({
                        "Tanggal": str(tanggal_absen),
                        "Sekolah": pilih_sekolah_form,
                        "Mata_Pelajaran": mata_pelajaran,
                        "Kelas": pilih_kelas_form,
                        "No": int(no_s) if pd.notna(no_s) else idx + 1,
                        "Nama_Siswa": nama_s,
                        "S": "V" if s_checked else "",
                        "I": "V" if i_checked else "",
                        "A": "V" if a_checked else ""
                    })
                    
                if st.button("💾 Simpan Absensi ke Database Pribadi"):
                    if absensi_rows:
                        for row in absensi_rows:
                            sheet_absensi.append_row([
                                row["Tanggal"], row["Sekolah"], row["Mata_Pelajaran"], 
                                row["Kelas"], row["No"], row["Nama_Siswa"], 
                                row["S"], row["I"], row["A"]
                            ])
                        st.success("Data absensi berhasil disimpan ke Spreadsheet pribadi Anda!")
                        st.rerun()
                    else:
                        st.warning("Tidak ada data siswa untuk disimpan.")
            else:
                st.info("Belum ada data siswa pada kelas ini.")

    with tab2:
        st.subheader("📥 Download & Upload Database Guru")
        st.markdown("Silakan unduh template resmi yang mencakup 3 sheet (Absensi Harian, Rekap Semester Ganjil, Rekap Semester Genap) langsung dari Google Sheets.")
        
        # ID Google Sheets Template Database Guru
        template_guru_id = "1F2o-ODBFDHezkhSZas5jtaD6vM8Ny8KF2gZ4KtVTM7w"
        url_download_template = f"https://docs.google.com/spreadsheets/d/{template_guru_id}/export?format=xlsx"
        st.write(f"DEBUG URL: {url_download_template}")
        
        st.markdown(
            f"""
            <a href="{url_download_template}" target="_blank">
                <button style="background-color: #ff4b4b; color: white; padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 14px;">
                    📥 Download Template Database Guru (.XLSX)
                </button>
            </a>
            """,
            unsafe_allow_html=True
        )
        
        st.markdown("---")
        st.markdown("### Upload File Excel yang Telah Diedit")
        uploaded_file = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])
        if uploaded_file is not None:
            if st.button("📤 Proses & Sinkronkan"):
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    if "Absensi Harian" in xls.sheet_names:
                        df_up_abs = pd.read_excel(xls, "Absensi Harian")
                        sheet_absensi.clear()
                        sheet_absensi.update([df_up_abs.columns.values.tolist()] + df_up_abs.values.tolist())
                    if "Rekap Semester Ganjil" in xls.sheet_names:
                        df_up_ganjil = pd.read_excel(xls, "Rekap Semester Ganjil")
                        sheet_rekap_ganjil.clear()
                        sheet_rekap_ganjil.update([df_up_ganjil.columns.values.tolist()] + df_up_ganjil.values.tolist())
                    if "Rekap Semester Genap" in xls.sheet_names:
                        df_up_genap = pd.read_excel(xls, "Rekap Semester Genap")
                        sheet_rekap_genap.clear()
                        sheet_rekap_genap.update([df_up_genap.columns.values.tolist()] + df_up_genap.values.tolist())
                    st.success("Sinkronisasi database berhasil!")
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")

    with tab3:
        st.subheader("📈 Laporan Harian Absensi")
        if not df_absensi.empty:
            filtered_df = df_absensi.copy()
            if pilih_sekolah != "Semua Sekolah" and "Sekolah" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["Sekolah"] == pilih_sekolah]
            if pilih_kelas != "Semua Kelas" and "Kelas" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["Kelas"] == pilih_kelas]
            st.dataframe(filtered_df, use_container_width=True)
        else:
            st.info("Belum ada data absensi harian.")

    with tab4:
        st.subheader("📊 Rekap Semester Ganjil")
        try:
            data_ganjil = sheet_rekap_ganjil.get_all_records()
            df_ganjil = pd.DataFrame(data_ganjil)
            if not df_ganjil.empty:
                if pilih_sekolah != "Semua Sekolah" and "Nama_Sekolah" in df_ganjil.columns:
                    df_ganjil = df_ganjil[df_ganjil["Nama_Sekolah"] == pilih_sekolah]
                if pilih_kelas != "Semua Kelas" and "Kelas" in df_ganjil.columns:
                    df_ganjil = df_ganjil[df_ganjil["Kelas"] == pilih_kelas]
                st.dataframe(df_ganjil, use_container_width=True)
            else:
                st.info("Belum ada data Rekap Semester Ganjil.")
        except Exception as e:
            st.error(f"Gagal memuat rekap ganjil: {e}")

    with tab5:
        st.subheader("📊 Rekap Semester Genap")
        try:
            data_genap = sheet_rekap_genap.get_all_records()
            df_genap = pd.DataFrame(data_genap)
            if not df_genap.empty:
                if pilih_sekolah != "Semua Sekolah" and "Nama_Sekolah" in df_genap.columns:
                    df_genap = df_genap[df_genap["Nama_Sekolah"] == pilih_sekolah]
                if pilih_kelas != "Semua Kelas" and "Kelas" in df_genap.columns:
                    df_genap = df_genap[df_genap["Kelas"] == pilih_kelas]
                st.dataframe(df_genap, use_container_width=True)
            else:
                st.info("Belum ada data Rekap Semester Genap.")
        except Exception as e:
            st.error(f"Gagal memuat rekap genap: {e}")