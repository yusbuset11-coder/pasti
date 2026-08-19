import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gspread_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def verifikasi_login(email, token):
    try:
        client = get_gspread_client()
        master_sheet_id = "1mgN63xzrLt__5b9-gBw8dIWYP3RRgNdagUiTurFZdgg"
        sheet = client.open_by_key(master_sheet_id).sheet1
        data = sheet.get_all_records()

        if not data:
            st.warning("⚠️ Spreadsheet Master Registry kosong atau tidak terbaca.")
            return None

        for row in data:
            db_email = str(row.get("Email", "")).strip().lower()
            db_token = str(row.get("Token_Unik", "")).strip()
            db_status = str(row.get("Status", "")).strip().upper()

            if db_email == email.strip().lower() and db_token == token.strip():
                if db_status == "AKTIF":
                    return row
                else:
                    st.error("❌ Akun Anda berstatus TIDAK AKTIF.")
                    return None
                    
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan koneksi ke database pusat: {e}")
    return None

def render_halaman_login():
    st.markdown("## 🔐 Login Aplikasi PASTI")
    st.write("Silakan masukkan Email dan Token Unik Anda untuk mengakses sistem.")

    with st.form("form_login"):
        email_input = st.text_input("Email Terdaftar")
        token_input = st.text_input("Token Unik", type="password")
        tombol_login = st.form_submit_button("Masuk")

        if tombol_login:
            if not email_input or not token_input:
                st.warning("⚠️ Email dan Token Unik wajib diisi!")
            else:
                user_data = verifikasi_login(email_input, token_input)
                if user_data:
                    st.session_state["logged_in"] = True
                    st.session_state["nama_guru"] = user_data.get("Nama_Guru")
                    st.session_state["email_guru"] = user_data.get("Email")
                    st.session_state["sheet_id_guru"] = user_data.get("Spreadsheet_ID_Guru")
                    st.success(f"✅ Login berhasil! Selamat datang, {user_data['Nama_Guru']}")
                    st.rerun()
                else:
                    st.error("❌ Login gagal. Email atau Token Unik salah.")