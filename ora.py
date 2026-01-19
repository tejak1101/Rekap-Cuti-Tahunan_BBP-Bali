# import streamlit as st
# import pandas as pd
# import plotly.express as px
# import re
# from io import BytesIO

# # ================= CONFIG =================
# st.set_page_config(
#     page_title="Dashboard Cuti Pegawai",
#     page_icon="📊",
#     layout="wide"
# )

# # ================= STYLE =================
# st.markdown("""
# <style>
# body {
#     background: linear-gradient(120deg,#0f172a,#020617);
# }
# h1,h2,h3 {color:#f8fafc;}
# </style>
# """, unsafe_allow_html=True)

# # ================= HEADER =================
# st.title("📊 Dashboard Cuti Tahunan 2025")
# st.subheader("Balai Bahasa Provinsi Bali")
# st.markdown("---")

# # ================= UPLOAD =================
# file = st.file_uploader(
#     "📤 Upload File Excel Cuti Pegawai",
#     type=["xlsx"]
# )

# if file:
#     df = pd.read_excel(file)

#     # Bersihkan nama kolom
#     df.columns = df.columns.str.strip().str.lower()

#     st.info(f"Kolom terdeteksi: {list(df.columns)}")

#     # ================= AUTO DETEKSI KOLOM LAMA =================
#     lama_col = None
#     for col in df.columns:
#         if "lama" in col:
#             lama_col = col

#     if not lama_col:
#         st.error("❌ Kolom 'Lama' tidak ditemukan")
#         st.stop()

#     # ================= CLEAN DATA =================
#     df[lama_col] = (
#         df[lama_col]
#         .astype(str)
#         .apply(lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
#     )

#     # ================= REKAP =================
#     rekap = df.groupby(['nip','nama pegawai']).agg(
#         jumlah_pengajuan=(lama_col,'count'),
#         total_hari_cuti=(lama_col,'sum')
#     ).reset_index()

#     rekap['status'] = rekap['total_hari_cuti'].apply(
#         lambda x: '⚠ Melebihi Batas' if x > 12 else '✅ Aman'
#     )

#     # ================= KPI =================
#     c1,c2,c3,c4 = st.columns(4)

#     c1.metric("👥 Total Pegawai", rekap.shape[0])
#     c2.metric("📆 Total Hari Cuti", int(rekap['total_hari_cuti'].sum()))
#     c3.metric("📊 Rata-rata Cuti", round(rekap['total_hari_cuti'].mean(),1))
#     c4.metric("⚠ Over Limit", rekap[rekap['total_hari_cuti']>12].shape[0])

#     st.markdown("---")

#     # ================= TABS =================
#     tab1,tab2,tab3,tab4 = st.tabs([
#         "📌 Rekap Pegawai",
#         "🏆 Ranking",
#         "📈 Tren Bulanan",
#         "⚠ Warning"
#     ])

#     # ---- TAB 1
#     with tab1:
#         st.dataframe(rekap, use_container_width=True)

#     # ---- TAB 2
#     with tab2:
#         top5 = rekap.sort_values(
#             'total_hari_cuti',ascending=False
#         ).head(5)

#         fig = px.bar(
#             top5,
#             x='nama pegawai',
#             y='total_hari_cuti',
#             text='total_hari_cuti',
#             color='total_hari_cuti',
#             template="plotly_dark"
#         )

#         st.plotly_chart(fig, use_container_width=True)

#     # ---- TAB 3
#     with tab3:
#         # auto deteksi kolom tanggal
#         tgl_col = None
#         for col in df.columns:
#             if "mulai" in col or "tanggal" in col:
#                 tgl_col = col

#         if tgl_col:
#             df['bulan'] = pd.to_datetime(df[tgl_col]).dt.month_name()

#             bulanan = df.groupby('bulan')[lama_col].sum().reset_index()

#             fig2 = px.line(
#                 bulanan,
#                 x='bulan',
#                 y=lama_col,
#                 markers=True,
#                 template="plotly_dark"
#             )

#             st.plotly_chart(fig2, use_container_width=True)
#         else:
#             st.warning("Kolom tanggal tidak ditemukan")

#     # ---- TAB 4
#     with tab4:
#         warning = rekap[rekap['total_hari_cuti']>=10]
#         st.dataframe(warning, use_container_width=True)

#     # ================= GRAFIK UTAMA =================
#     st.subheader("📊 Perbandingan Semua Pegawai")

#     fig_all = px.bar(
#         rekap,
#         x='nama pegawai',
#         y='total_hari_cuti',
#         color='status',
#         template="plotly_dark"
#     )

#     st.plotly_chart(fig_all, use_container_width=True)

#     # ================= DOWNLOAD =================
#     buffer = BytesIO()
#     rekap.to_excel(buffer,index=False)
#     buffer.seek(0)

#     st.download_button(
#         "📥 Download Excel Rekap",
#         buffer,
#         file_name="Rekap_Cuti_2025.xlsx"
#     )

#     st.success("Dashboard berhasil digenerate 🎉")

# else:
#     st.info("Silakan upload file Excel dulu")

# import streamlit as st
# import pandas as pd
# import plotly.express as px
# import re
# from io import BytesIO

# # ================= CONFIG =================
# st.set_page_config(
#     page_title="Dashboard Cuti Pegawai",
#     page_icon="📊",
#     layout="wide"
# )

# # ================= STYLE =================
# st.markdown("""
# <style>
# body {
#     background: linear-gradient(135deg,#020617,#0f172a);
# }
# h1,h2,h3 {color:#f8fafc;}
# </style>
# """, unsafe_allow_html=True)

# # ================= HEADER =================
# st.title("📊 Dashboard Cuti Tahunan 2025")
# st.subheader("Balai Bahasa Provinsi Bali")
# st.markdown("---")

# # ================= UPLOAD =================
# file = st.file_uploader(
#     "📤 Upload File Excel Cuti Pegawai",
#     type=["xlsx"]
# )

# if file:
#     df = pd.read_excel(file)

#     # Bersihkan nama kolom
#     df.columns = df.columns.str.strip().str.lower()

#     # AUTO DETEKSI KOLOM LAMA
#     lama_col = None
#     for col in df.columns:
#         if "lama" in col:
#             lama_col = col

#     if not lama_col:
#         st.error("❌ Kolom 'Lama' tidak ditemukan")
#         st.stop()

#     # CLEAN DATA
#     df[lama_col] = (
#         df[lama_col]
#         .astype(str)
#         .apply(lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
#     )

#     # ================= REKAP =================
#     rekap = df.groupby(['nip','nama pegawai']).agg(
#         jumlah_pengajuan=(lama_col,'count'),
#         total_hari_cuti=(lama_col,'sum')
#     ).reset_index()

#     rekap['status'] = rekap['total_hari_cuti'].apply(
#         lambda x: '⚠ Melebihi Batas' if x > 12 else '✅ Aman'
#     )

#     # ================= KPI =================
#     c1,c2,c3,c4 = st.columns(4)

#     c1.metric("👥 Total Pegawai", rekap.shape[0])
#     c2.metric("📆 Total Hari Cuti", int(rekap['total_hari_cuti'].sum()))
#     c3.metric("📊 Rata-rata Cuti", round(rekap['total_hari_cuti'].mean(),1))
#     c4.metric("⚠ Over Limit", rekap[rekap['total_hari_cuti']>12].shape[0])

#     st.markdown("---")

#     # ================= TABS =================
#     tab1,tab2,tab3,tab4 = st.tabs([
#         "📌 Rekap Pegawai",
#         "🏆 Ranking",
#         "📈 Tren Bulanan",
#         "⚠ Warning"
#     ])

#     # ---- TAB 1
#     with tab1:
#         st.dataframe(rekap, use_container_width=True)

#     # ---- TAB 2
#     with tab2:
#         top5 = rekap.sort_values(
#             'total_hari_cuti',ascending=False
#         ).head(5)

#         fig_rank = px.bar(
#             top5,
#             x='nama pegawai',
#             y='total_hari_cuti',
#             text='total_hari_cuti',
#             color='total_hari_cuti',
#             template="plotly_dark",
#             title="🏆 Top 5 Pegawai Cuti Terbanyak"
#         )

#         fig_rank.update_traces(
#             textposition='outside',
#             marker_line_width=1.5
#         )

#         fig_rank.update_layout(
#             transition_duration=900
#         )

#         st.plotly_chart(fig_rank, use_container_width=True)

#     # ---- TAB 3
#     with tab3:
#         tgl_col = None
#         for col in df.columns:
#             if "mulai" in col or "tanggal" in col:
#                 tgl_col = col

#         if tgl_col:
#             df['bulan'] = pd.to_datetime(df[tgl_col]).dt.month_name()

#             bulanan = df.groupby('bulan')[lama_col].sum().reset_index()

#             fig_month = px.line(
#                 bulanan,
#                 x='bulan',
#                 y=lama_col,
#                 markers=True,
#                 template="plotly_dark",
#                 title="📈 Tren Cuti Per Bulan"
#             )

#             fig_month.update_layout(
#                 transition_duration=900
#             )

#             st.plotly_chart(fig_month, use_container_width=True)
#         else:
#             st.warning("Kolom tanggal tidak ditemukan")

#     # ---- TAB 4
#     with tab4:
#         warning = rekap[rekap['total_hari_cuti']>=10]
#         st.dataframe(warning, use_container_width=True)

#     # ================= GRAFIK UTAMA =================
#     st.subheader("📊 Perbandingan Semua Pegawai")

#     # SORT KECIL → BESAR
#     rekap_sorted = rekap.sort_values(
#         by='total_hari_cuti',
#         ascending=True
#     )

#     fig_all = px.bar(
#         rekap_sorted,
#         x='nama pegawai',
#         y='total_hari_cuti',
#         color='status',
#         text='total_hari_cuti',
#         template="plotly_dark"
#     )

#     fig_all.update_traces(
#         textposition='outside'
#     )

#     fig_all.update_layout(
#         title="Perbandingan Cuti Pegawai (Sedikit → Banyak)",
#         xaxis_tickangle=-45,
#         height=600,
#         transition_duration=1200
#     )

#     st.plotly_chart(fig_all, use_container_width=True)

#     # ================= DOWNLOAD =================
#     buffer = BytesIO()
#     rekap.to_excel(buffer,index=False)
#     buffer.seek(0)

#     st.download_button(
#         "📥 Download Excel Rekap",
#         buffer,
#         file_name="Rekap_Cuti_2025.xlsx"
#     )

#     st.success("Dashboard berhasil digenerate 🎉")

# else:
#     st.info("Silakan upload file Excel dulu")







# import streamlit as st
# import pandas as pd
# import plotly.express as px
# import re
# from io import BytesIO

# # ================= CONFIG =================
# st.set_page_config(
#     page_title="Dashboard Cuti Pegawai",
#     page_icon="📊",
#     layout="wide"
# )

# # ================= LUXURY STYLE =================
# st.markdown("""
# <style>
# body {
#     background: linear-gradient(135deg,#020617,#0f172a,#020617);
# }
# h1,h2,h3 {color:#f8fafc;}
# .stTabs [data-baseweb="tab"] {
#     font-size:16px;
#     font-weight:600;
# }
# </style>
# """, unsafe_allow_html=True)

# # ================= HEADER =================
# st.title("📊 Dashboard Cuti Tahunan 2025")
# st.subheader("Balai Bahasa Provinsi Bali")
# st.markdown("---")

# # ================= UPLOAD =================
# file = st.file_uploader(
#     "📤 Upload File Excel Cuti Pegawai",
#     type=["xlsx"]
# )

# if file:
#     df = pd.read_excel(file)

#     # Bersihkan kolom
#     df.columns = df.columns.str.strip().str.lower()

#     # AUTO DETEKSI KOLOM LAMA
#     lama_col = [c for c in df.columns if "lama" in c][0]

#     # CLEAN
#     df[lama_col] = (
#         df[lama_col]
#         .astype(str)
#         .apply(lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
#     )

#     # ================= REKAP UTAMA =================
#     rekap = df.groupby(['nip','nama pegawai']).agg(
#         jumlah_pengajuan=(lama_col,'count'),
#         total_hari_cuti=(lama_col,'sum')
#     ).reset_index()

#     rekap['status'] = rekap['total_hari_cuti'].apply(
#         lambda x: '⚠ Melebihi Batas' if x > 12 else '✅ Aman'
#     )

#     # ================= KPI =================
#     c1,c2,c3,c4 = st.columns(4)
#     c1.metric("👥 Total Pegawai", rekap.shape[0])
#     c2.metric("📆 Total Hari Cuti", int(rekap['total_hari_cuti'].sum()))
#     c3.metric("📊 Rata-rata", round(rekap['total_hari_cuti'].mean(),1))
#     c4.metric("⚠ Over Limit", rekap[rekap['total_hari_cuti']>12].shape[0])

#     st.markdown("---")

#     # ================= TABS =================
#     tab1,tab2,tab3,tab4 = st.tabs([
#         "📌 Rekap Pegawai",
#         "🏆 Ranking",
#         "📈 Tren Bulanan",
#         "⚠ Warning"
#     ])

#     # ---------- TAB 1 ----------
#     with tab1:
#         st.dataframe(rekap, use_container_width=True)

#     # ---------- TAB 2 ----------
#     with tab2:
#         ranking = rekap.sort_values(
#             'total_hari_cuti',ascending=False
#         )

#         fig_rank = px.bar(
#             ranking.head(10),
#             x='nama pegawai',
#             y='total_hari_cuti',
#             color='total_hari_cuti',
#             color_continuous_scale='Turbo',
#             text='total_hari_cuti',
#             template="plotly_dark"
#         )

#         fig_rank.update_layout(
#             title="🏆 Ranking Pegawai",
#             transition_duration=1200
#         )

#         st.plotly_chart(fig_rank, use_container_width=True)

#     # ---------- TAB 3 ----------
#     with tab3:
#         tgl_col = [c for c in df.columns if "mulai" in c or "tanggal" in c]

#         if tgl_col:
#             df['bulan'] = pd.to_datetime(df[tgl_col[0]]).dt.month_name()

#             bulanan = df.groupby('bulan')[lama_col].sum().reset_index()

#             fig_month = px.line(
#                 bulanan,
#                 x='bulan',
#                 y=lama_col,
#                 markers=True,
#                 template="plotly_dark",
#                 color_discrete_sequence=["#22c55e"]
#             )

#             fig_month.update_layout(
#                 title="📈 Tren Cuti Bulanan",
#                 transition_duration=1200
#             )

#             st.plotly_chart(fig_month, use_container_width=True)
#         else:
#             st.warning("Kolom tanggal tidak ditemukan")

#     # ---------- TAB 4 ----------
#     with tab4:
#         warning = rekap[rekap['total_hari_cuti']>=10]
#         st.dataframe(warning, use_container_width=True)

#     # ================= GRAFIK UTAMA =================
#     st.subheader("📊 Perbandingan Pegawai")

#     rekap_sorted = rekap.sort_values(
#         'total_hari_cuti',ascending=True
#     )

#     fig_all = px.bar(
#         rekap_sorted,
#         x='nama pegawai',
#         y='total_hari_cuti',
#         color='total_hari_cuti',
#         color_continuous_scale='Plasma',
#         text='total_hari_cuti',
#         template="plotly_dark"
#     )

#     fig_all.update_layout(
#         title="Sedikit → Banyak",
#         transition_duration=1500
#     )

#     st.plotly_chart(fig_all, use_container_width=True)

#     # ================= EXPORT MULTI SHEET =================
#     buffer = BytesIO()
#     with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
#         rekap.to_excel(writer, sheet_name="Rekap Pegawai", index=False)
#         ranking.to_excel(writer, sheet_name="Ranking", index=False)
#         bulanan.to_excel(writer, sheet_name="Tren Bulanan", index=False)
#         warning.to_excel(writer, sheet_name="Warning", index=False)
#         df.to_excel(writer, sheet_name="Data Mentah", index=False)

#     buffer.seek(0)

#     st.download_button(
#         "📥 Download Excel Lengkap (Multi Sheet)",
#         buffer,
#         file_name="Dashboard_Cuti_2025_LUXURY.xlsx"
#     )

#     st.success("✨ Dashboard luxury + Excel multi sheet siap!")

# else:
#     st.info("Silakan upload file Excel dulu")








import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from io import BytesIO

# ================= CONFIG =================
st.set_page_config(
    page_title="Dashboard Cuti Pegawai",
    page_icon="📊",
    layout="wide"
)

# ================= ULTRA LUXURY STYLE =================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap');

* {
    font-family: 'Inter', sans-serif;
}

[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0a0e27 0%, #1a1f3a 25%, #0f1123 50%, #1e1b4b 75%, #0a0e27 100%);
    background-size: 400% 400%;
    animation: gradientShift 15s ease infinite;
}

@keyframes gradientShift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

[data-testid="stHeader"] {
    background: transparent;
}

h1 {
    background: linear-gradient(135deg, #60a5fa 0%, #a78bfa 50%, #f472b6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 800;
    font-size: 3.5rem !important;
    text-align: center;
    margin-bottom: 0.5rem;
    letter-spacing: -1px;
    animation: glow 3s ease-in-out infinite;
}

@keyframes glow {
    0%, 100% { filter: drop-shadow(0 0 20px rgba(96, 165, 250, 0.5)); }
    50% { filter: drop-shadow(0 0 40px rgba(167, 139, 250, 0.8)); }
}

h2 {
    color: #e0e7ff !important;
    font-weight: 600;
    font-size: 1.8rem !important;
    margin-top: 2rem;
}

h3 {
    color: #c7d2fe !important;
    font-weight: 600;
}

.stSubheader {
    text-align: center;
    color: #a5b4fc !important;
    font-size: 1.3rem;
    font-weight: 400;
    margin-bottom: 2rem;
}

/* Metric Cards */
[data-testid="metric-container"] {
    background: linear-gradient(135deg, rgba(30, 41, 59, 0.6) 0%, rgba(51, 65, 85, 0.4) 100%);
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 20px;
    padding: 1.5rem;
    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
    backdrop-filter: blur(10px);
    transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

[data-testid="metric-container"]:hover {
    transform: translateY(-8px);
    box-shadow: 0 30px 80px rgba(96, 165, 250, 0.3);
    border-color: rgba(96, 165, 250, 0.5);
}

[data-testid="metric-container"] label {
    color: #94a3b8 !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 1px;
}

[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #f1f5f9 !important;
    font-size: 2.5rem !important;
    font-weight: 700 !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: rgba(15, 23, 42, 0.6);
    padding: 8px;
    border-radius: 16px;
    backdrop-filter: blur(10px);
}

.stTabs [data-baseweb="tab"] {
    background: transparent;
    color: #94a3b8;
    border-radius: 12px;
    font-size: 16px;
    font-weight: 600;
    padding: 12px 24px;
    transition: all 0.3s ease;
}

.stTabs [data-baseweb="tab"]:hover {
    background: rgba(96, 165, 250, 0.1);
    color: #60a5fa;
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #3b82f6, #8b5cf6) !important;
    color: white !important;
    box-shadow: 0 8px 24px rgba(59, 130, 246, 0.4);
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.4);
}

/* Upload */
[data-testid="stFileUploader"] {
    background: linear-gradient(135deg, rgba(30, 41, 59, 0.5), rgba(51, 65, 85, 0.3));
    border: 2px dashed rgba(96, 165, 250, 0.4);
    border-radius: 20px;
    padding: 2rem;
    transition: all 0.3s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: rgba(96, 165, 250, 0.8);
    background: rgba(59, 130, 246, 0.1);
}

/* Download Button */
.stDownloadButton button {
    background: linear-gradient(135deg, #10b981, #059669) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 12px 32px !important;
    font-weight: 600 !important;
    font-size: 16px !important;
    box-shadow: 0 10px 30px rgba(16, 185, 129, 0.4) !important;
    transition: all 0.3s ease !important;
}

.stDownloadButton button:hover {
    transform: translateY(-3px);
    box-shadow: 0 15px 40px rgba(16, 185, 129, 0.6) !important;
}

/* Divider */
hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent, rgba(96, 165, 250, 0.5), transparent);
    margin: 2rem 0;
}

/* Success/Info/Warning */
.stSuccess, .stInfo, .stWarning {
    background: rgba(30, 41, 59, 0.6) !important;
    border-radius: 12px !important;
    border-left: 4px solid !important;
    backdrop-filter: blur(10px);
}

.stSuccess {
    border-left-color: #10b981 !important;
}

.stInfo {
    border-left-color: #3b82f6 !important;
}

.stWarning {
    border-left-color: #f59e0b !important;
}

/* Logo Container */
.logo-container {
    text-align: center;
    margin-bottom: 2rem;
    animation: float 3s ease-in-out infinite;
}

@keyframes float {
    0%, 100% { transform: translateY(0px); }
    50% { transform: translateY(-10px); }
}

.logo-img {
    max-width: 200px;
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(96, 165, 250, 0.3);
    background: white;
    padding: 15px;
}

/* Select Box */
.stSelectbox [data-baseweb="select"] {
    background: rgba(30, 41, 59, 0.6);
    border-radius: 12px;
}

</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.markdown(f'''
<div class="logo-container">
    <img src="https://internal-portal.kemdikbud.go.id/web/image/res.company/1/logo/unique_id" 
         class="logo-img" 
         onerror="this.style.display='none'; this.nextElementSibling.style.display='block';">
    <div style="display:none; font-size:4rem;">📊</div>
</div>
''', unsafe_allow_html=True)

st.title("Dashboard Cuti Tahunan")
st.subheader("Balai Bahasa Provinsi Bali")
st.markdown("---")

# ================= UPLOAD =================
file = st.file_uploader(
    "📤 Upload File Excel Cuti Pegawai",
    type=["xlsx"]
)

if file:
    # Baca semua sheet
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    
    st.info(f"📋 Ditemukan {len(sheet_names)} sheet: {', '.join(sheet_names)}")
    
    # Pilih sheet untuk diproses
    selected_sheet = st.selectbox(
        "🔍 Pilih Sheet untuk Dianalisis:",
        sheet_names,
        index=0 if len(sheet_names) > 0 else None
    )
    
    if selected_sheet:
        df = pd.read_excel(file, sheet_name=selected_sheet)
        
        # Tampilkan info kolom
        with st.expander("🔍 Lihat Struktur Data"):
            st.write(f"**Jumlah Baris:** {len(df)}")
            st.write(f"**Kolom yang tersedia:** {', '.join(df.columns.tolist())}")
            st.dataframe(df.head(), use_container_width=True)
        
        # Bersihkan kolom
        df.columns = df.columns.str.strip().str.lower()
        
        # AUTO DETEKSI KOLOM
        lama_col = None
        nip_col = None
        nama_col = None
        tgl_col = None
        
        for c in df.columns:
            if "lama" in c:
                lama_col = c
            if "nip" in c:
                nip_col = c
            if "nama" in c and "pegawai" in c:
                nama_col = c
            if "mulai" in c or ("tanggal" in c and "mulai" in c):
                tgl_col = c
        
        # Validasi kolom penting
        if not lama_col:
            st.error("❌ Kolom 'Lama Cuti' tidak ditemukan!")
            st.stop()
        
        if not nip_col:
            st.warning("⚠️ Kolom NIP tidak ditemukan, akan menggunakan index")
            df['nip'] = df.index
            nip_col = 'nip'
        
        if not nama_col:
            st.error("❌ Kolom 'Nama Pegawai' tidak ditemukan!")
            st.stop()
        
        # CLEAN data lama cuti
        df[lama_col] = (
            df[lama_col]
            .astype(str)
            .apply(lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
        )
        
        # ================= REKAP UTAMA =================
        rekap = df.groupby([nip_col, nama_col]).agg(
            jumlah_pengajuan=(lama_col, 'count'),
            total_hari_cuti=(lama_col, 'sum')
        ).reset_index()
        
        rekap['status'] = rekap['total_hari_cuti'].apply(
            lambda x: '⚠ Melebihi Batas' if x > 12 else '✅ Aman'
        )
        
        # ================= KPI =================
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("👥 Total Pegawai", rekap.shape[0])
        c2.metric("📆 Total Hari Cuti", int(rekap['total_hari_cuti'].sum()))
        c3.metric("📊 Rata-rata", round(rekap['total_hari_cuti'].mean(), 1))
        c4.metric("⚠ Over Limit", rekap[rekap['total_hari_cuti'] > 12].shape[0])
        
        st.markdown("---")
        
        # ================= TABS =================
        tab1, tab2, tab3, tab4 = st.tabs([
            "📌 Rekap Pegawai",
            "🏆 Ranking",
            "📈 Tren Bulanan",
            "⚠ Warning"
        ])
        
        # ---------- TAB 1 ----------
        with tab1:
            st.dataframe(rekap, use_container_width=True)
        
        # ---------- TAB 2 ----------
        with tab2:
            ranking = rekap.sort_values('total_hari_cuti', ascending=False)
            
            fig_rank = go.Figure(data=[
                go.Bar(
                    x=ranking.head(10)[nama_col],
                    y=ranking.head(10)['total_hari_cuti'],
                    text=ranking.head(10)['total_hari_cuti'],
                    textposition='outside',
                    marker=dict(
                        color=ranking.head(10)['total_hari_cuti'],
                        colorscale='Turbo',
                        line=dict(color='rgba(255,255,255,0.3)', width=2),
                        opacity=0.9
                    ),
                    hovertemplate='<b>%{x}</b><br>Total: %{y} hari<extra></extra>'
                )
            ])
            
            fig_rank.update_layout(
                title=dict(text="🏆 Top 10 Ranking Pegawai", font=dict(size=24, color='#e0e7ff')),
                template="plotly_dark",
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Inter', color='#cbd5e1'),
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor='rgba(148,163,184,0.1)'),
                height=500
            )
            
            st.plotly_chart(fig_rank, use_container_width=True)
        
        # ---------- TAB 3 ----------
        with tab3:
            if tgl_col:
                df['bulan'] = pd.to_datetime(df[tgl_col], errors='coerce').dt.month_name()
                
                bulanan = df.groupby('bulan')[lama_col].sum().reset_index()
                
                # Urutkan bulan
                bulan_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                              'July', 'August', 'September', 'October', 'November', 'December']
                bulanan['bulan'] = pd.Categorical(bulanan['bulan'], categories=bulan_order, ordered=True)
                bulanan = bulanan.sort_values('bulan')
                
                fig_month = go.Figure(data=[
                    go.Scatter(
                        x=bulanan['bulan'],
                        y=bulanan[lama_col],
                        mode='lines+markers',
                        line=dict(color='#10b981', width=4, shape='spline'),
                        marker=dict(size=12, color='#10b981', line=dict(color='white', width=2)),
                        fill='tozeroy',
                        fillcolor='rgba(16, 185, 129, 0.2)',
                        hovertemplate='<b>%{x}</b><br>Total: %{y} hari<extra></extra>'
                    )
                ])
                
                fig_month.update_layout(
                    title=dict(text="📈 Tren Cuti Bulanan", font=dict(size=24, color='#e0e7ff')),
                    template="plotly_dark",
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family='Inter', color='#cbd5e1'),
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='rgba(148,163,184,0.1)'),
                    height=500
                )
                
                st.plotly_chart(fig_month, use_container_width=True)
            else:
                st.warning("⚠️ Kolom tanggal tidak ditemukan untuk analisis tren bulanan")
        
        # ---------- TAB 4 ----------
        with tab4:
            warning = rekap[rekap['total_hari_cuti'] >= 10]
            if len(warning) > 0:
                st.dataframe(warning, use_container_width=True)
            else:
                st.success("✅ Tidak ada pegawai yang mendekati atau melebihi batas cuti!")
        
        # ================= GRAFIK UTAMA =================
        st.subheader("📊 Perbandingan Pegawai")
        
        rekap_sorted = rekap.sort_values('total_hari_cuti', ascending=True)
        
        fig_all = go.Figure(data=[
            go.Bar(
                x=rekap_sorted[nama_col],
                y=rekap_sorted['total_hari_cuti'],
                text=rekap_sorted['total_hari_cuti'],
                textposition='outside',
                marker=dict(
                    color=rekap_sorted['total_hari_cuti'],
                    colorscale='Plasma',
                    line=dict(color='rgba(255,255,255,0.2)', width=1.5),
                    opacity=0.95
                ),
                hovertemplate='<b>%{x}</b><br>Total: %{y} hari<extra></extra>'
            )
        ])
        
        fig_all.update_layout(
            title=dict(text="Distribusi Cuti: Sedikit → Banyak", font=dict(size=24, color='#e0e7ff')),
            template="plotly_dark",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Inter', color='#cbd5e1'),
            xaxis=dict(showgrid=False, tickangle=-45),
            yaxis=dict(showgrid=True, gridcolor='rgba(148,163,184,0.1)'),
            height=600
        )
        
        st.plotly_chart(fig_all, use_container_width=True)
        
        # ================= EXPORT MULTI SHEET =================
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            rekap.to_excel(writer, sheet_name="Rekap Pegawai", index=False)
            ranking.to_excel(writer, sheet_name="Ranking", index=False)
            if tgl_col:
                bulanan.to_excel(writer, sheet_name="Tren Bulanan", index=False)
            warning.to_excel(writer, sheet_name="Warning", index=False)
            df.to_excel(writer, sheet_name="Data Mentah", index=False)
        
        buffer.seek(0)
        
        st.download_button(
            "📥 Download Excel Lengkap (Multi Sheet)",
            buffer,
            file_name=f"Dashboard_Cuti_{selected_sheet}.xlsx"
        )
        
        st.success(f"✨ Dashboard untuk sheet '{selected_sheet}' berhasil dibuat!")

else:
    st.info("📁 Silakan upload file Excel untuk memulai analisis")