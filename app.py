
import os, shutil, base64
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "data" / "database.xlsx"
LOGO_PATH = APP_DIR / "assets" / "logo_al_ghozali.jpg"

st.set_page_config(
    page_title="Sistem PKG Guru Al-Ghozali",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "menu" not in st.session_state:
    st.session_state.menu = "Dashboard"

def logo_data_uri():
    if not LOGO_PATH.exists():
        return ""
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode()
    return f"data:image/jpeg;base64,{encoded}"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');
:root {
    --navy:#062b55;
    --blue:#0b67b2;
    --blue2:#208fe5;
    --gold:#d5a52c;
    --gold2:#f0c95c;
    --cream:#fffaf0;
    --ink:#17304f;
}
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: linear-gradient(135deg,#f7fbff 0%,#fffdf7 55%,#f1f7ff 100%); }
[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 1.2rem; max-width: 1450px; }
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#062b55 0%,#084477 55%,#052a50 100%);
    border-right: 1px solid rgba(255,255,255,.12);
}
[data-testid="stSidebar"] * { color: #fff !important; }
[data-testid="stSidebar"] .stButton > button {
    border-radius: 999px !important;
    border: 1px solid rgba(255,255,255,.16) !important;
    background: rgba(255,255,255,.08) !important;
    color: white !important;
    text-align: left !important;
    margin: 3px 0 !important;
    transition: .2s ease;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: linear-gradient(90deg,var(--gold),#e6bc48) !important;
    color: #09284b !important;
    transform: translateX(3px);
}
.brand { text-align:center; padding: 8px 8px 20px; }
.brand img { width: 112px; height:112px; object-fit:contain; border-radius:50%; background:white; padding:6px; box-shadow:0 8px 24px rgba(0,0,0,.2); }
.brand h2 { font-family:'Plus Jakarta Sans'; font-size:19px; margin:10px 0 2px; color:#f3cc63; }
.brand p { font-size:11px; color:#dbeaff !important; margin:0; }
.gold-line { height:3px; width:80px; background:linear-gradient(90deg,transparent,#f0c95c,transparent); margin:14px auto; }
.hero {
    background: linear-gradient(110deg,#062b55 0%,#0a4c83 62%,#d5a52c 160%);
    border-radius: 26px; padding: 28px 32px; color:white;
    box-shadow:0 15px 38px rgba(6,43,85,.18); position:relative; overflow:hidden;
}
.hero:after { content:""; position:absolute; right:-80px; top:-100px; width:310px; height:310px; border:2px solid rgba(255,255,255,.12); border-radius:50%; box-shadow:0 0 0 25px rgba(255,255,255,.05),0 0 0 55px rgba(255,255,255,.035); }
.hero h1 { font-family:'Plus Jakarta Sans'; font-size:31px; margin:0 0 8px; }
.hero p { margin:0; opacity:.9; }
.section-title { color:var(--navy); font-family:'Plus Jakarta Sans'; font-size:21px; font-weight:800; margin:24px 0 12px; }
.kpi {
    background:white; border:1px solid #e6edf5; border-radius:22px; padding:18px;
    box-shadow:0 8px 25px rgba(8,52,93,.07); min-height:120px;
}
.kpi .label { font-size:12px; color:#6d7f92; font-weight:700; }
.kpi .value { font-size:30px; font-weight:800; color:var(--navy); margin-top:8px; }
.kpi.gold { border-top:5px solid var(--gold); }
.kpi.blue { border-top:5px solid var(--blue2); }
.kpi small { color:#75889d; }
.card {
    background:white; border:1px solid #e6edf5; border-radius:22px; padding:20px;
    box-shadow:0 8px 25px rgba(8,52,93,.06);
}
div[data-testid="stMetric"] { background:transparent; }
.stButton > button {
    border-radius:999px; font-weight:700; border:1px solid #d8e3ee;
}
div[data-testid="stForm"] { background:#fff; border:1px solid #e6edf5; border-radius:20px; padding:18px; }
.stDownloadButton > button { border-radius:999px; background:var(--navy); color:white; }
[data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)

def sheets():
    return pd.ExcelFile(DB_PATH).sheet_names

def raw_sheet(name):
    return pd.read_excel(DB_PATH, sheet_name=name, header=None)

def table_sheet(name):
    raw = raw_sheet(name)
    header_idx = 0
    for i in range(min(15, len(raw))):
        if raw.iloc[i].notna().sum() >= 2:
            header_idx = i
            break
    df = pd.read_excel(DB_PATH, sheet_name=name, header=header_idx)
    return df.dropna(how="all").reset_index(drop=True)

def save_sheet(name, df):
    xls = pd.ExcelFile(DB_PATH)
    temp = DB_PATH.parent / "database_temp.xlsx"
    backup = DB_PATH.parent / "database_backup.xlsx"
    if temp.exists():
        temp.unlink()
    with pd.ExcelWriter(temp, engine="openpyxl") as writer:
        for s in xls.sheet_names:
            if s == name:
                df.to_excel(writer, sheet_name=s, index=False)
            else:
                pd.read_excel(DB_PATH, sheet_name=s, header=None).to_excel(
                    writer, sheet_name=s, index=False, header=False
                )
    shutil.copy2(DB_PATH, backup)
    os.replace(str(temp), str(DB_PATH))

def teacher_names():
    if "Data Guru" not in sheets():
        return []
    df = table_sheet("Data Guru")
    col = next((c for c in df.columns if "nama guru" in str(c).lower()), None)
    return sorted(df[col].dropna().astype(str).unique().tolist()) if col else []

def pill_menu(label, icon):
    return f"{icon}  {label}"

# Sidebar
with st.sidebar:
    logo = logo_data_uri()
    st.markdown(f"""
    <div class="brand">
      <img src="{logo}">
      <h2>YAYASAN PENDIDIKAN ISLAM PONDOK MODERN AL-GHOZALI</h2>
      <p>Sistem Penilaian Kinerja Guru</p>
      <div class="gold-line"></div>
    </div>
    """, unsafe_allow_html=True)
    menu = st.radio(
        "NAVIGASI",
        ["Dashboard","Data Guru","Administrasi Guru","Penilaian Kompetensi","Prestasi Kerja","Kehadiran","Rekap Nilai","Cetak Rapor","Master Database"],
        key="menu",
        format_func=lambda x: {
            "Dashboard":"⌂  Dashboard",
            "Data Guru":"♙  Data Guru",
            "Administrasi Guru":"📁  Administrasi Guru",
            "Penilaian Kompetensi":"▣  Penilaian Kompetensi",
            "Prestasi Kerja":"🏆  Prestasi Kerja",
            "Kehadiran":"◷  Kehadiran",
            "Rekap Nilai":"▤  Rekap Nilai",
            "Cetak Rapor":"▤  Cetak Rapor",
            "Master Database":"◉  Master Database",
        }[x]
    )
    st.markdown("---")
    st.caption("DATABASE UTAMA")
    st.success("Spreadsheet terhubung")
    st.caption("© Yayasan Pendidikan Islam Pondok Modern Al-Ghozali")

# Tombol navigasi kembali ke dashboard
if st.session_state.get("menu", "Dashboard") != "Dashboard":
    nav_col, _ = st.columns([1, 5])
    with nav_col:
        if st.button("← Kembali ke Dashboard", key="back_dashboard", use_container_width=True):
            st.session_state.menu = "Dashboard"
            st.rerun()

# Header
st.markdown(f"""
<div class="hero">
  <h1>SISTEM PKG GURU</h1>
  <p>Yayasan Pendidikan Islam Pondok Modern Al-Ghozali · Profesional · Akuntabel · Transparan · Berkelanjutan</p>
</div>
""", unsafe_allow_html=True)

if menu == "Dashboard":
    st.markdown('<div class="section-title">Selamat Datang di Dashboard PKG</div>', unsafe_allow_html=True)
    dg = table_sheet("Data Guru") if "Data Guru" in sheets() else pd.DataFrame()
    rk = table_sheet("Rekap Nilai") if "Rekap Nilai" in sheets() else pd.DataFrame()
    cols = st.columns(4)
    metrics = [
        ("Jumlah Guru", len(dg), "gold", "Data master guru"),
        ("Data Rekap", len(rk), "blue", "Baris rekap penilaian"),
        ("Sheet Database", len(sheets()), "gold", "Terhubung spreadsheet"),
        ("Status Sistem", "Aktif", "blue", "Siap digunakan"),
    ]
    for col, (label, value, style, sub) in zip(cols, metrics):
        with col:
            st.markdown(f'<div class="kpi {style}"><div class="label">{label}</div><div class="value">{value}</div><small>{sub}</small></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Menu Utama</div>', unsafe_allow_html=True)
    menu_cols = st.columns(4)
    shortcuts = [("Data Guru","👥"),("Administrasi Guru","📁"),("Penilaian Kompetensi","▣"),("Prestasi Kerja","🏆"),("Kehadiran","◷"),("Rekap Nilai","▤"),("Cetak Rapor","▤"),("Master Database","◉")]
    for i, (name, icon) in enumerate(shortcuts):
        with menu_cols[i % 4]:
            st.markdown(f'<div class="card" style="text-align:center;margin-bottom:14px"><div style="font-size:30px">{icon}</div><b style="color:#062b55">{name}</b></div>', unsafe_allow_html=True)
    left, right = st.columns([1.4,1])
    with left:
        st.markdown('<div class="section-title">Ringkasan Data Guru</div>', unsafe_allow_html=True)
        st.dataframe(dg.head(20), use_container_width=True, hide_index=True)
    with right:
        st.markdown('<div class="section-title">Distribusi Unit</div>', unsafe_allow_html=True)
        if not dg.empty and "Unit" in dg.columns:
            counts = dg["Unit"].fillna("Tidak Diisi").value_counts()
            st.bar_chart(counts)
        else:
            st.info("Belum ada data unit.")

elif menu == "Data Guru":
    st.markdown('<div class="section-title">Data Guru</div>', unsafe_allow_html=True)
    df = table_sheet("Data Guru")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.markdown('<div class="section-title">Tambah Data Guru</div>', unsafe_allow_html=True)
    with st.form("tambah_guru"):
        vals = {str(c): st.text_input(str(c)) for c in df.columns}
        ok = st.form_submit_button("Simpan Data Guru")
    if ok:
        save_sheet("Data Guru", pd.concat([df, pd.DataFrame([vals])], ignore_index=True))
        st.success("Data guru berhasil disimpan ke spreadsheet.")

elif menu in ["Penilaian Kompetensi","Prestasi Kerja","Kehadiran"]:
    st.markdown(f'<div class="section-title">{menu}</div>', unsafe_allow_html=True)
    df = table_sheet(menu)

    nama_guru_col = next(
        (c for c in df.columns if str(c).strip().lower() in ["nama guru", "nama_guru", "guru", "nama"]),
        None
    )
    guru_terpilih = st.selectbox(
        "Pilih Nama Guru",
        ["-- Pilih Guru --"] + teacher_names(),
        key=f"guru_{menu}"
    )

    if guru_terpilih == "-- Pilih Guru --":
        st.info("Silakan pilih nama guru terlebih dahulu. Data nilai akan ditampilkan khusus untuk guru yang dipilih.")
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    else:
        if nama_guru_col is not None:
            df_guru = df[
                df[nama_guru_col].fillna("").astype(str).str.strip().str.casefold()
                == guru_terpilih.strip().casefold()
            ].copy()
        else:
            df_guru = df.iloc[0:0].copy()

        st.markdown(
            f'<div class="card"><b>Guru yang sedang dinilai:</b> '
            f'<span style="color:#d5a52c">{guru_terpilih}</span></div>',
            unsafe_allow_html=True
        )
        st.write("### Data Nilai Guru Terpilih")
        if df_guru.empty:
            st.info("Belum ada data penilaian untuk guru ini.")
        else:
            st.dataframe(df_guru, use_container_width=True, hide_index=True)

        st.write("### Input Nilai")
        with st.form(f"input_penilaian_{menu}"):
            vals = {}
            for c in df.columns[:12]:
                if c == nama_guru_col:
                    st.text_input(str(c), value=guru_terpilih, disabled=True)
                    vals[str(c)] = guru_terpilih
                else:
                    vals[str(c)] = st.text_input(str(c))
            ok = st.form_submit_button("Simpan Nilai Guru Ini")
        if ok:
            if nama_guru_col is not None:
                vals[str(nama_guru_col)] = guru_terpilih
            save_sheet(menu, pd.concat([df, pd.DataFrame([vals])], ignore_index=True))
            st.success(f"Nilai {guru_terpilih} berhasil disimpan ke spreadsheet.")
            st.rerun()

elif menu == "Rekap Nilai":
    st.markdown('<div class="section-title">Rekap Nilai Kinerja Guru</div>', unsafe_allow_html=True)
    df = table_sheet("Rekap Nilai")
    st.dataframe(df, use_container_width=True, hide_index=True)
    numeric = df.select_dtypes(include="number")
    if not numeric.empty:
        st.markdown('<div class="section-title">Visualisasi Nilai</div>', unsafe_allow_html=True)
        st.bar_chart(numeric.mean())

elif menu == "Cetak Rapor":
    st.markdown('<div class="section-title">Cetak Rapor Kinerja Guru</div>', unsafe_allow_html=True)
    teacher = st.selectbox("Pilih Guru", ["-- Pilih Guru --"] + teacher_names())
    if teacher != "-- Pilih Guru --":
        df_guru = table_sheet("Data Guru") if "Data Guru" in sheets() else pd.DataFrame()
        df_rekap = table_sheet("Rekap Nilai") if "Rekap Nilai" in sheets() else pd.DataFrame()
        
        nama_col_guru = next((c for c in df_guru.columns if str(c).strip().lower() in ["nama guru", "nama_guru", "guru", "nama"]), None)
        nama_col_rekap = next((c for c in df_rekap.columns if str(c).strip().lower() in ["nama guru", "nama_guru", "guru", "nama"]), None)
        
        guru_info = {}
        if nama_col_guru is not None and not df_guru.empty:
            baris = df_guru[df_guru[nama_col_guru].fillna("").astype(str).str.strip().str.casefold() == teacher.strip().casefold()]
            if not baris.empty:
                guru_info = baris.iloc[0].to_dict()
                
        rekap_info = {}
        if nama_col_rekap is not None and not df_rekap.empty:
            baris = df_rekap[df_rekap[nama_col_rekap].fillna("").astype(str).str.strip().str.casefold() == teacher.strip().casefold()]
            if not baris.empty:
                rekap_info = baris.iloc[0].to_dict()
                
        st.markdown("---")
        
        # Inject CSS khusus untuk mode Print (Ctrl+P)
        st.markdown("""
        <style>
        @media print {
            body * { visibility: hidden; }
            #rapor-area, #rapor-area * { visibility: visible; }
            #rapor-area {
                position: absolute; left: 0; top: 0; width: 100%;
                box-shadow: none !important; border: none !important;
            }
            [data-testid="stSidebar"], [data-testid="stHeader"] { display: none !important; }
        }
        </style>
        """, unsafe_allow_html=True)
        
        rapor_html = f"""
        <div id="rapor-area" style="background:white; padding: 40px; border: 1px solid #e6edf5; border-radius: 12px; margin-bottom: 20px; color: black; font-family: 'Times New Roman', Times, serif; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
            <div style="text-align: center; border-bottom: 3px solid black; padding-bottom: 15px; margin-bottom: 25px;">
                <h2 style="margin: 0; font-size: 24px; text-transform: uppercase;">YAYASAN PENDIDIKAN ISLAM PONDOK MODERN AL-GHOZALI</h2>
                <h3 style="margin: 8px 0 0; font-size: 18px;">LAPORAN PENILAIAN KINERJA GURU (PKG)</h3>
            </div>
            <table style="width: 100%; margin-bottom: 25px; font-size: 15px; border: none;">
                <tr><td style="width: 150px; font-weight: bold; border: none; padding: 4px 0;">Nama Guru</td><td style="border: none; padding: 4px 0;">: {teacher}</td></tr>
                <tr><td style="font-weight: bold; border: none; padding: 4px 0;">ID Guru / NUPTK</td><td style="border: none; padding: 4px 0;">: {guru_info.get('ID Guru', guru_info.get('NUPTK', '-'))}</td></tr>
                <tr><td style="font-weight: bold; border: none; padding: 4px 0;">Unit Kerja</td><td style="border: none; padding: 4px 0;">: {guru_info.get('Unit', '-')}</td></tr>
            </table>
            <h4 style="margin-bottom: 12px; font-size: 16px;">Rincian Hasil Penilaian:</h4>
            <table style="width: 100%; border-collapse: collapse; font-size: 15px; text-align: left; margin-bottom: 30px;">
                <thead>
                    <tr style="background-color: #f9fafc;">
                        <th style="border: 1px solid black; padding: 10px;">Kriteria / Indikator Penilaian</th>
                        <th style="border: 1px solid black; padding: 10px; text-align: center; width: 120px;">Nilai</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for k, v in rekap_info.items():
            if str(k).strip().lower() not in ["nama guru", "nama_guru", "guru", "nama", "id guru", "nuptk", "unit"]:
                if pd.notna(v) and str(v).strip() != "":
                    val_str = str(v)
                    if isinstance(v, (int, float)):
                        val_str = f"{v:.2f}".rstrip('0').rstrip('.') if '.' in f"{v:.2f}" else str(v)
                    
                    rapor_html += f"""
                    <tr>
                        <td style="border: 1px solid black; padding: 10px;">{k}</td>
                        <td style="border: 1px solid black; padding: 10px; text-align: center; font-weight: bold;">{val_str}</td>
                    </tr>
                    """
                    
        rapor_html += f"""
                </tbody>
            </table>
            <div style="margin-top: 50px; width: 100%; display: flex; justify-content: space-between;">
                <div style="text-align: center; width: 45%;">
                    <p style="margin: 0;">Mengetahui,</p>
                    <p style="margin: 0; margin-bottom: 80px;">Kepala Sekolah</p>
                    <p style="margin: 0; text-decoration: underline; font-weight: bold;">( .................................... )</p>
                </div>
                <div style="text-align: center; width: 45%;">
                    <p style="margin: 0;">Guru yang Dinilai,</p>
                    <p style="margin: 0; margin-bottom: 80px;"><br></p>
                    <p style="margin: 0; text-decoration: underline; font-weight: bold;">{teacher}</p>
                </div>
            </div>
        </div>
        """
        
        # Gunakan st.html jika tersedia (Streamlit >= 1.34), atau st.markdown
        if hasattr(st, 'html'):
            st.html(rapor_html)
        else:
            st.markdown(rapor_html, unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 4])
        with col1:
            st.download_button(
                label="⬇️ Download Rapor (HTML)",
                data=rapor_html,
                file_name=f"Rapor_PKG_{teacher.replace(' ', '_')}.html",
                mime="text/html",
                type="primary",
                use_container_width=True
            )
        with col2:
            st.caption("File akan diunduh dalam format HTML. Anda dapat mengkliknya dua kali untuk membukanya di browser, lalu mencetaknya/simpan ke PDF.")

elif menu == "Master Database":
    st.markdown('<div class="section-title">Master Database Spreadsheet</div>', unsafe_allow_html=True)
    for name in ["MASTER UNIT","MASTER MAPEL","MASTER GURU","MASTER KELAS","PENUGASAN GURU"]:
        if name in sheets():
            with st.expander(name, expanded=(name=="MASTER GURU")):
                st.dataframe(table_sheet(name), use_container_width=True, hide_index=True)

elif menu == "Administrasi Guru":
    st.markdown('<div class="section-title">Deteksi File Administrasi Guru</div>', unsafe_allow_html=True)
    st.info("Fitur ini mendeteksi file administrasi pada folder lokal (Komputer/Flashdisk) Anda secara real-time. Jika nama file mengandung nama guru, statusnya akan otomatis berubah menjadi 'Sudah'.")
    
    folder_path = st.text_input(
        "Path Folder Administrasi (Lokal)", 
        value=st.session_state.get("admin_folder", ""),
        placeholder="Contoh: D:\\Administrasi_Guru"
    )
    
    if folder_path:
        st.session_state["admin_folder"] = folder_path
        if os.path.isdir(folder_path):
            st.success(f"Folder ditemukan: `{folder_path}`")
            files = os.listdir(folder_path)
            teachers = teacher_names()
            
            if not teachers:
                st.warning("Belum ada data guru di sistem.")
            else:
                status_data = []
                for t in teachers:
                    t_normalized = str(t).lower().replace(" ", "").replace("_", "")
                    
                    found_files = []
                    for f in files:
                        f_normalized = str(f).lower().replace(" ", "").replace("_", "")
                        if t_normalized in f_normalized:
                            found_files.append(f)
                            
                    status_data.append({
                        "Nama Guru": t,
                        "Status": "✅ Sudah" if found_files else "❌ Belum",
                        "File Ditemukan": ", ".join(found_files) if found_files else "-"
                    })
                
                df_status = pd.DataFrame(status_data)
                
                def highlight_status(val):
                    color = '#e6ffe6' if '✅' in str(val) else '#ffe6e6'
                    return f'background-color: {color}'
                
                st.dataframe(
                    df_status.style.map(highlight_status, subset=['Status']),
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.error("Folder tidak ditemukan. Pastikan path yang dimasukkan benar dan aplikasi memiliki izin akses ke folder tersebut.")
