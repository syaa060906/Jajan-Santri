import streamlit as st
import pandas as pd
import datetime
from google.oauth2.service_account import Credentials
import gspread

# --- CONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Pencatat Jajan Santri",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SINKRONISASI GOOGLE SHEETS ---
@st.cache_resource
def init_connection():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )
    return gspread.authorize(credentials)

try:
    gc = init_connection()
    sh = gc.open("ADM JAJAN SANTRI")
except Exception as e:
    st.error("Belum terhubung ke Google Sheets Secrets Streamlit. Harap atur Secrets terlebih dahulu.")
    st.stop()

# --- HELPER FUNCTIONS ---
def load_data():
    ws_rekap = sh.worksheet("Lembar2")
    data = ws_rekap.get_all_values()
    
    df = pd.DataFrame(data[4:], columns=data[3]) # skip header
    df = df.dropna(how='all')
    df = df.iloc[:, 1:9]
    df.columns = ['No', 'Nama Santri', 'L/P', 'Kelas', 'Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo']
    df = df[df['Nama Santri'] != '']
    
    for col in ['Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo']:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace('.', '').str.replace(',', ''), errors='coerce').fillna(0)
        
    return df

def tambah_santri_baru(nama, lp, kelas, jatah_jajan):
    """Menambahkan santri baru ke Sheet Lembar2 & Membuat Sheet Riwayat Baru"""
    ws_rekap = sh.worksheet("Lembar2")
    all_vals = ws_rekap.get_all_values()
    
    # Hitung No Urut Baru
    last_row = len(all_vals) + 1
    no_baru = len(all_vals) - 4 + 1
    
    # Append row baru ke Lembar2
    ws_rekap.append_row([
        "", no_baru, nama.upper(), lp, kelas, jatah_jajan, 0, 0, 0
    ])
    
    # Buat Sheet Baru Khusus Santri Tersebut
    try:
        ws_new = sh.add_worksheet(title=str(no_baru), rows="100", cols="10")
        ws_new.append_row(["", "REKAP UANG JAJAN SANTRI"])
        ws_new.append_row(["", "BMT (BAITUL MAL WA TAMWIL)"])
        ws_new.append_row(["", "PONDOK PESANTREN MIFTAHUL HUDA IV"])
        ws_new.append_row(["", "Nama", "", nama.upper()])
        ws_new.append_row(["", "Kelas", "", kelas])
        ws_new.append_row(["", "Jatah perhari", "", jatah_jajan])
        ws_new.append_row(["", "No.", "Tanggal", "Keterangan", "Uang Masuk", "Uang Keluar", "Saldo"])
    except:
        pass

def catat_transaksi(no_santri, jenis, nominal, tanggal, keterangan):
    ws_rekap = sh.worksheet("Lembar2")
    row_idx = int(no_santri) + 4
    
    val_masuk = float(ws_rekap.cell(row_idx, 7).value or 0)
    val_keluar = float(ws_rekap.cell(row_idx, 8).value or 0)
    
    if "Masuk" in jenis:
        new_masuk = val_masuk + nominal
        ws_rekap.update_cell(row_idx, 7, new_masuk)
    else:
        new_keluar = val_keluar + nominal
        ws_rekap.update_cell(row_idx, 8, new_keluar)
        
    current_masuk = float(ws_rekap.cell(row_idx, 7).value or 0)
    current_keluar = float(ws_rekap.cell(row_idx, 8).value or 0)
    ws_rekap.update_cell(row_idx, 9, current_masuk - current_keluar)

    try:
        ws_santri = sh.worksheet(str(no_santri))
        next_row = len(ws_santri.get_all_values()) + 1
        no_trx = next_row - 7
        
        uang_masuk = nominal if "Masuk" in jenis else ""
        uang_keluar = nominal if "Keluar" in jenis else ""
        saldo_baru = current_masuk - current_keluar
        
        ws_santri.append_row([
            "", no_trx, str(tanggal), keterangan, uang_masuk, uang_keluar, saldo_baru
        ])
    except:
        pass
    return True

# --- HEADER APP ---
st.title("💰 System Keuangan Jajan Santri")
st.caption("Pondok Pesantren Miftahul Huda IV - Cloud Access")

df_santri = load_data()

if not df_santri.empty:
    # --- DASHBOARD ---
    st.subheader("📊 Summary Keuangan Total")
    col1, col2, col3 = st.columns(3)
    
    total_masuk = df_santri['Uang Masuk'].sum()
    total_keluar = df_santri['Uang Keluar'].sum()
    total_saldo = df_santri['Sisa Saldo'].sum()
    
    col1.metric("Total Uang Masuk", f"Rp {total_masuk:,.0f}".replace(",", "."))
    col2.metric("Total Terpakai (Jajan)", f"Rp {total_keluar:,.0f}".replace(",", "."))
    col3.metric("Total Sisa Saldo", f"Rp {total_saldo:,.0f}".replace(",", "."))
    
    st.divider()

    # --- SIDEBAR MENU ---
    st.sidebar.header("📌 Menu Navigasi")
    menu = st.sidebar.radio("Pilih Halaman:", [
        "Input Transaksi Cepat", 
        "➕ Tambah Santri Baru", 
        "Daftar Saldo Santri"
    ])

    # -------------------------------------------------------------
    # MENU 1: INPUT TRANSAKSI
    # -------------------------------------------------------------
    if menu == "Input Transaksi Cepat":
        st.subheader("📝 Catat Transaksi Baru")
        with st.form("form_transaksi", clear_on_submit=True):
            nama_santri = st.selectbox("Pilih Nama Santri:", df_santri['Nama Santri'].unique())
            
            c1, c2 = st.columns(2)
            with c1:
                jenis_trx = st.selectbox("Jenis Transaksi:", ["Uang Keluar (Jajan/Beli)", "Uang Masuk (Kiriman Ortu)"])
            with c2:
                nominal = st.number_input("Nominal (Rp):", min_value=500, step=1000, value=10000)
                
            tgl = st.date_input("Tanggal:", datetime.date.today())
            keterangan = st.text_input("Keterangan:", value="Jajan Harian" if "Keluar" in jenis_trx else "Kiriman Ortu")
            
            submitted = st.form_submit_button("💾 Simpan Ke Cloud")
            if submitted:
                no_santri = df_santri[df_santri['Nama Santri'] == nama_santri]['No'].values[0]
                if catat_transaksi(no_santri, jenis_trx, nominal, tgl, keterangan):
                    st.success(f"Berhasil menyimpan data untuk **{nama_santri}**!")
                    st.rerun()

    # -------------------------------------------------------------
    # MENU 2: TAMBAH SANTRI BARU
    # -------------------------------------------------------------
    elif menu == "➕ Tambah Santri Baru":
        st.subheader("➕ Tambah Santri Baru ke System")
        with st.form("form_tambah_santri", clear_on_submit=True):
            nama_baru = st.text_input("Nama Lengkap Santri:")
            c_lp, c_kelas = st.columns(2)
            with c_lp:
                lp_baru = st.selectbox("Jenis Kelamin (L/P):", ["L", "P"])
            with c_kelas:
                kelas_baru = st.text_input("Kelas (misal: 7 SMP, 10 SMK):", value="7 SMP")
                
            jatah_baru = st.number_input("Jatah Jajan Perhari (Rp):", min_value=0, step=5000, value=20000)
            
            btn_tambah = st.form_submit_button("💾 Tambahkan Santri")
            if btn_tambah:
                if nama_baru.strip() != "":
                    tambah_santri_baru(nama_baru, lp_baru, kelas_baru, jatah_baru)
                    st.success(f"Santri **{nama_baru}** berhasil ditambahkan ke Google Sheets!")
                    st.rerun()
                else:
                    st.warning("Nama santri tidak boleh kosong!")

    # -------------------------------------------------------------
    # MENU 3: DAFTAR SALDO
    # -------------------------------------------------------------
    elif menu == "Daftar Saldo Santri":
        st.subheader("📋 Rekapitulasi Data Santri")
        st.dataframe(
            df_santri[['No', 'Nama Santri', 'Kelas', 'Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo']],
            use_container_width=True,
            hide_index=True
        )
