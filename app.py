import streamlit as st
import pandas as pd
import datetime
import re
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

# --- HELPER CLEANER ANGKA ---
def clean_number(value):
    if not value:
        return 0.0
    val_str = str(value)
    cleaned = re.sub(r'[^\d]', '', val_str)
    return float(cleaned) if cleaned else 0.0

# --- LOAD DATA REKAP ---
def load_data():
    ws_rekap = sh.worksheet("Lembar2")
    data = ws_rekap.get_all_values()
    
    # Ambil dari baris 6 ke bawah
    rows = data[5:]
    
    parsed_data = []
    for r in rows:
        # Col B (idx 1) = No, Col C (idx 2) = Nama
        if len(r) > 2 and r[2].strip() != "":
            no = r[1] if len(r) > 1 else ""
            nama = r[2]
            lp = r[3] if len(r) > 3 else ""
            kelas = r[4] if len(r) > 4 else ""
            
            # Pas dengan kolom F, G, H, I (idx 5, 6, 7, 8)
            jajan = clean_number(r[5]) if len(r) > 5 else 0.0
            masuk = clean_number(r[6]) if len(r) > 6 else 0.0
            keluar = clean_number(r[7]) if len(r) > 7 else 0.0
            saldo = clean_number(r[8]) if len(r) > 8 else 0.0
            
            parsed_data.append([no, nama, lp, kelas, jajan, masuk, keluar, saldo])
            
    df = pd.DataFrame(parsed_data, columns=['No', 'Nama Santri', 'L/P', 'Kelas', 'Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo'])
    return df

# --- RIWAYAT PER SANTRI ---
def get_riwayat_santri(no_santri):
    try:
        ws_santri = sh.worksheet(str(no_santri))
        data = ws_santri.get_all_values()
        
        if len(data) >= 7:
            rows = data[6:]
            parsed_trx = []
            for r in rows:
                if len(r) > 3 and (r[3].strip() != "" or r[2].strip() != ""):
                    no_trx = r[1] if len(r) > 1 else ""
                    tgl = r[2] if len(r) > 2 else ""
                    ket = r[3] if len(r) > 3 else ""
                    masuk = clean_number(r[4]) if len(r) > 4 else 0.0
                    keluar = clean_number(r[5]) if len(r) > 5 else 0.0
                    saldo = clean_number(r[6]) if len(r) > 6 else 0.0
                    
                    parsed_trx.append([no_trx, tgl, ket, masuk, keluar, saldo])
            
            return pd.DataFrame(parsed_trx, columns=['No. Trx', 'Tanggal', 'Keterangan', 'Uang Masuk', 'Uang Keluar', 'Saldo Sisa'])
    except Exception:
        pass
    return pd.DataFrame()

# --- TAMBAH SANTRI BARU ---
def tambah_santri_baru(nama, lp, kelas, jatah_jajan):
    ws_rekap = sh.worksheet("Lembar2")
    all_vals = ws_rekap.get_all_values()
    
    target_row = None
    existing_count = 0
    
    for i in range(5, len(all_vals)):
        row = all_vals[i]
        nama_cell = row[2].strip() if len(row) > 2 else ""
        if nama_cell != "":
            existing_count += 1
        elif target_row is None:
            target_row = i + 1
            
    if target_row is None:
        target_row = len(all_vals) + 1
        
    no_baru = existing_count + 1
    
    # TEPAT PAS 8 KOLOM (Col B sampai Col I):
    # [No, Nama, L/P, Kelas, Jajan, Uang Masuk, Uang Keluar, Sisa Saldo]
    data_row = [int(no_baru), nama.upper(), lp, kelas, int(jatah_jajan), 0, 0, 0]
    ws_rekap.update(f"B{target_row}:I{target_row}", [data_row])
    
    # Buat Sheet Individual untuk Santri Baru
    try:
        ws_new = sh.add_worksheet(title=str(no_baru), rows="100", cols="10")
        ws_new.append_row(["", "REKAP UANG JAJAN SANTRI"])
        ws_new.append_row(["", "BMT (BAITUL MAL WA TAMWIL)"])
        ws_new.append_row(["", "PONDOK PESANTREN MIFTAHUL HUDA IV"])
        ws_new.append_row(["", "Nama", "", nama.upper()])
        ws_new.append_row(["", "Kelas", "", kelas])
        ws_new.append_row(["", "Jatah perhari", "", jatah_jajan])
        ws_new.append_row(["", "No.", "Tanggal", "Keterangan", "Uang Masuk", "Uang Keluar", "Saldo"])
    except Exception:
        pass

# --- CATAT TRANSAKSI ---
def catat_transaksi(no_santri, jenis, nominal, tanggal, keterangan):
    ws_rekap = sh.worksheet("Lembar2")
    
    # Cari baris santri berdasarkan No di Kolom B
    cell = ws_rekap.find(str(no_santri), in_column=2)
    if not cell:
        return False
        
    row_idx = cell.row
    
    # Col G (7) = Uang Masuk | Col H (8) = Uang Keluar | Col I (9) = Sisa Saldo
    val_masuk = clean_number(ws_rekap.cell(row_idx, 7).value)
    val_keluar = clean_number(ws_rekap.cell(row_idx, 8).value)
    
    if "Masuk" in jenis:
        new_masuk = val_masuk + nominal
        ws_rekap.update_cell(row_idx, 7, int(new_masuk))
    else:
        new_keluar = val_keluar + nominal
        ws_rekap.update_cell(row_idx, 8, int(new_keluar))
        
    current_masuk = clean_number(ws_rekap.cell(row_idx, 7).value)
    current_keluar = clean_number(ws_rekap.cell(row_idx, 8).value)
    
    # Update Sisa Saldo di Col I
    ws_rekap.update_cell(row_idx, 9, int(current_masuk - current_keluar))

    # Catat di Sheet Santri Individual
    try:
        ws_santri = sh.worksheet(str(no_santri))
        next_row = len(ws_santri.get_all_values()) + 1
        no_trx = next_row - 6
        
        uang_masuk = int(nominal) if "Masuk" in jenis else ""
        uang_keluar = int(nominal) if "Keluar" in jenis else ""
        saldo_baru = int(current_masuk - current_keluar)
        
        ws_santri.append_row([
            "", no_trx, str(tanggal), keterangan, uang_masuk, uang_keluar, saldo_baru
        ])
    except Exception:
        pass
    return True

# --- HAPUS SANTRI ---
def hapus_santri(nama_target):
    ws_rekap = sh.worksheet("Lembar2")
    cell = ws_rekap.find(nama_target, in_column=3)
    
    if cell:
        row_idx = cell.row
        no_santri = ws_rekap.cell(row_idx, 2).value
        
        empty_row = ["", "", "", "", "", "", "", ""]
        ws_rekap.update(f"B{row_idx}:I{row_idx}", [empty_row])
        
        try:
            ws_santri = sh.worksheet(str(no_santri))
            sh.del_worksheet(ws_santri)
        except Exception:
            pass
        return True
    return False

# --- UI STREAMLIT ---
st.title("💰 System Keuangan Jajan Santri")
st.caption("Pondok Pesantren Miftahul Huda IV - Cloud Access")

if "msg_success" in st.session_state:
    st.success(st.session_state.msg_success)
    del st.session_state.msg_success

df_santri = load_data()

if not df_santri.empty:
    st.subheader("📊 Summary Keuangan Total")
    col1, col2, col3 = st.columns(3)
    
    total_masuk = df_santri['Uang Masuk'].sum()
    total_keluar = df_santri['Uang Keluar'].sum()
    total_saldo = df_santri['Sisa Saldo'].sum()
    
    col1.metric("Total Uang Masuk", f"Rp {total_masuk:,.0f}".replace(",", "."))
    col2.metric("Total Terpakai (Jajan)", f"Rp {total_keluar:,.0f}".replace(",", "."))
    col3.metric("Total Sisa Saldo", f"Rp {total_saldo:,.0f}".replace(",", "."))
    
    st.divider()

    st.sidebar.header("📌 Menu Navigasi")
    menu = st.sidebar.radio("Pilih Halaman:", [
        "Input Transaksi Cepat", 
        "📜 Riwayat & Mutasi Santri",
        "📋 Rekapitulasi Saldo",
        "➕ Tambah Santri Baru", 
        "❌ Hapus Data Santri"
    ])

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
                    st.session_state.msg_success = f"✅ Transaksi **{nama_santri}** berhasil disimpan!"
                    st.rerun()

    elif menu == "📜 Riwayat & Mutasi Santri":
        st.subheader("📜 Detail Riwayat Transaksi Santri")
        target_nama = st.selectbox("Pilih Nama Santri:", df_santri['Nama Santri'].unique())
        
        info_santri = df_santri[df_santri['Nama Santri'] == target_nama].iloc[0]
        no_santri = info_santri['No']
        
        c_i1, c_i2, c_i3, c_i4 = st.columns(4)
        c_i1.info(f"**Kelas:** {info_santri['Kelas']}")
        c_i2.info(f"**Jatah/Hari:** Rp {info_santri['Jatah Jajan']:,.0f}".replace(",", "."))
        c_i3.success(f"**Total Masuk:** Rp {info_santri['Uang Masuk']:,.0f}".replace(",", "."))
        c_i4.metric("Sisa Saldo", f"Rp {info_santri['Sisa Saldo']:,.0f}".replace(",", "."))
        
        st.markdown(f"### 📋 Catatan Mutasi: **{target_nama}**")
        df_riwayat = get_riwayat_santri(no_santri)
        
        if not df_riwayat.empty:
            df_display = df_riwayat.copy()
            df_display['Uang Masuk'] = df_display['Uang Masuk'].apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "-")
            df_display['Uang Keluar'] = df_display['Uang Keluar'].apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "-")
            df_display['Saldo Sisa'] = df_display['Saldo Sisa'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
            st.dataframe(df_display, use_container_width=True, hide_index=True)
        else:
            st.warning("Belum ada riwayat transaksi yang tercatat.")

    elif menu == "📋 Rekapitulasi Saldo":
        st.subheader("📋 Rekapitulasi Seluruh Santri")
        df_show = df_santri.copy()
        for c in ['Jatah Jajan', 'Uang Masuk', 'Uang Keluar', 'Sisa Saldo']:
            df_show[c] = df_show[c].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    elif menu == "➕ Tambah Santri Baru":
        st.subheader("➕ Tambah Santri Baru")
        with st.form("form_tambah_santri", clear_on_submit=True):
            nama_baru = st.text_input("Nama Lengkap Santri:")
            c_lp, c_kelas = st.columns(2)
            with c_lp:
                lp_baru = st.selectbox("Jenis Kelamin (L/P):", ["L", "P"])
            with c_kelas:
                kelas_baru = st.text_input("Kelas:", value="7 SMP")
                
            jatah_baru = st.number_input("Jatah Jajan Perhari (Rp):", min_value=0, step=5000, value=20000)
            
            btn_tambah = st.form_submit_button("💾 Tambahkan Santri")
            if btn_tambah:
                if nama_baru.strip() != "":
                    tambah_santri_baru(nama_baru, lp_baru, kelas_baru, jatah_baru)
                    st.session_state.msg_success = f"🎉 Data santri **{nama_baru.upper()}** berhasil ditambahkan!"
                    st.rerun()
                else:
                    st.warning("Nama santri tidak boleh kosong!")

    elif menu == "❌ Hapus Data Santri":
        st.subheader("❌ Hapus Data Santri")
        target_hapus = st.selectbox("Pilih Santri yang Akan Dihapus:", df_santri['Nama Santri'].unique())
        
        if st.button("🗑️ Hapus Santri Ini", type="primary"):
            if hapus_santri(target_hapus):
                st.session_state.msg_success = f"🗑️ Data santri **{target_hapus}** berhasil dihapus!"
                st.rerun()
            else:
                st.error("Gagal menghapus data santri.")
else:
    st.info("Belum ada data santri terisi.")
