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
        col1, col2 = st.columns(2)
        with col1:
            tanggal_absen = st.date_input("Tanggal Absensi", datetime.now())
            nama_guru_input = st.text_input("Nama Guru", value=user_name, disabled=True)
        with col2:
            sekolah_input = st.text_input("Nama Sekolah", value=pilih_sekolah if pilih_sekolah != "Semua Sekolah" else "")
            kelas_input = st.text_input("Kelas", value=pilih_kelas if pilih_kelas != "Semua Kelas" else "")
            
        mata_pelajaran = st.text_input("Mata Pelajaran", placeholder="Contoh: Pendidikan Pancasila")
        
        st.markdown("---")
        st.markdown("### Daftar Siswa & Status Kehadiran")
        
        jumlah_siswa = st.number_input("Jumlah Siswa di Kelas", min_value=1, max_value=50, value=5)
        
        absensi_rows = []
        for i in range(1, int(jumlah_siswa) + 1):
            cols = st.columns([1, 4, 1, 1, 1])
            with cols[0]:
                st.write(f"No. {i}")
            with cols[1]:
                nama_siswa = st.text_input(f"Nama Siswa {i}", key=f"siswa_{i}")
            with cols[2]:
                s = st.checkbox("S", key=f"s_{i}")
            with cols[3]:
                i_val = st.checkbox("I", key=f"i_{i}")
            with cols[4]:
                a = st.checkbox("A", key=f"a_{i}")
            
            if nama_siswa:
                absensi_rows.append({
                    "Tanggal": str(tanggal_absen),
                    "Sekolah": sekolah_input,
                    "Mata_Pelajaran": mata_pelajaran,
                    "Kelas": kelas_input,
                    "No": i,
                    "Nama_Siswa": nama_siswa,
                    "S": "V" if s else "",
                    "I": "V" if i_val else "",
                    "A": "V" if a else ""
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
                st.warning("Mohon isi minimal satu nama siswa.")

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