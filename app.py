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

gc = init_connection()
sh = gc.open_by_key("1RAIbKcuC1z1X_IYYS0yZzjdUp-qKL6uiEpO34INsW9g")

# --- HELPER FUNCTIONS ---
def load_data():
    ws_rekap = sh.worksheet("Lembar2")
    data = ws_rekap.get_all_values()
    
    rows = data[5:]  # Baris data mulai row 6 (index 5)
    
    parsed_data = []
    for r in rows:
        # Cek jika kolom C (Nama Santri, index 2) terisi
        if len(r) > 2 and r[2].strip() != "":
            no = r[1] if len(r) > 1 else ""
            nama = r[2]
            lp = r[3] if len(r) > 3 else ""
            kelas = r[4] if len(r) > 4 else ""
            jatah = r[6] if len(r) > 6 else "0"
            masuk = r[8] if len(r) > 8 else "0"
            keluar = r[10] if len(r) > 10 else "0"
            saldo = r[12] if len(r) > 12 else "0"
            
            parsed_data.append([no, nama, lp, kelas, jatah, masuk, keluar, saldo])
            
    df = pd.DataFrame(parsed_data, columns=['No', 'Nama Santri', 'L/P', 'Kelas', 'Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo'])
    
    for col in ['Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo']:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace('.', '').str.replace(',', '').str.replace('Rp', '').str.strip(), 
            errors='coerce'
        ).fillna(0)
        
    return df

def tambah_santri_baru(nama, lp, kelas, jatah_jajan):
    """Mengisi baris template kosong pertama di Lembar2 dari Kolom B s/d M"""
    ws_rekap = sh.worksheet("Lembar2")
    all_vals = ws_rekap.get_all_values()
    
    # Cari baris template pertama yang kolom Namanya (Kolom C) masih kosong
    target_row = None
    existing_count = 0
    
    for i in range(5, len(all_vals)):
        row = all_vals[i]
        nama_cell = row[2].strip() if len(row) > 2 else ""
        if nama_cell != "":
            existing_count += 1
        elif target_row is None:
            target_row = i + 1  # Row index 1-based di Sheets
            
    if target_row is None:
        target_row = len(all_vals) + 1
        
    no_baru = existing_count + 1
    
    # Update sel B{row} sampai M{row} secara presisi
    data_row = [no_baru, nama.upper(), lp, kelas, "Rp", jatah_jajan, "Rp", 0, "Rp", 0, "Rp", 0]
    ws_rekap.update(f"B{target_row}:M{target_row}", [data_row])
    
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
    cell = ws_rekap.find(str(no_santri), in_column=2)
    row_idx = cell.row
    
    val_masuk = float(str(ws_rekap.cell(row_idx, 9).value or 0).replace('.', '').replace(',', ''))
    val_keluar = float(str(ws_rekap.cell(row_idx, 11).value or 0).replace('.', '').replace(',', ''))
    
    if "Masuk" in jenis:
        new_masuk = val_masuk + nominal
        ws_rekap.update_cell(row_idx, 9, new_masuk)
    else:
        new_keluar = val_keluar + nominal
        ws_rekap.update_cell(row_idx, 11, new_keluar)
        
    current_masuk = float(str(ws_rekap.cell(row_idx, 9).value or 0).replace('.', '').replace(',', ''))
    current_keluar = float(str(ws_rekap.cell(row_idx, 11).value or 0).replace('.', '').replace(',', ''))
    ws_rekap.update_cell(row_idx, 13, current_masuk - current_keluar)

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

def hapus_santri(nama_target):
    """Menghapus data santri dari Lembar2 & menghapus Sheet riwayatnya"""
    ws_rekap = sh.worksheet("Lembar2")
    cell = ws_rekap.find(nama_target, in_column=3)
    
    if cell:
        row_idx = cell.row
        no_santri = ws_rekap.cell(row_idx, 2).value
        
        # Kosongkan baris B s/d M di Lembar2
        empty_row = ["", "", "", "", "", "", "", "", "", "", "", ""]
        ws_rekap.update(f"B{row_idx}:M{row_idx}", [empty_row])
        
        # Hapus Sheet Riwayat Individual jika ada
        try:
            ws_santri = sh.worksheet(str(no_santri))
            sh.del_worksheet(ws_santri)
        except:
            pass
        return True
    return False

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
        "Daftar Saldo Santri",
        "❌ Hapus Data Santri"
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
                    st.success(f"Santri **{nama_baru}** berhasil ditambahkan!")
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

    # -------------------------------------------------------------
    # MENU 4: HAPUS DATA SANTRI
    # -------------------------------------------------------------
    elif menu == "❌ Hapus Data Santri":
        st.subheader("❌ Hapus Data Santri dari System")
        st.warning("⚠️ Perhatian: Menghapus santri akan mengosongkan barisnya di rekap utama dan menghapus sheet riwayatnya.")
        
        target_hapus = st.selectbox("Pilih Santri yang Akan Dihapus:", df_santri['Nama Santri'].unique())
        
        if st.button("🗑️ Hapus Santri Ini", type="primary"):
            if hapus_santri(target_hapus):
                st.success(f"Data santri **{target_hapus}** berhasil dihapus!")
                st.rerun()
            else:
                st.error("Gagal menghapus data santri.")
else:
    st.info("Belum ada data santri terisi.")
