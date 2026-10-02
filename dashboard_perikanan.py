import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import html
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    Alignment,
    Border,
    Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

# ============================================================
# 1. KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(
    page_title="Dashboard Perikanan Indonesia",
    page_icon="🐟",
    layout="wide"
)

# ============================================================
# 2. CSS
# ============================================================
st.markdown("""
<style>
    /* -------------------------------------------------------
       SIDEBAR
       ------------------------------------------------------- */
    section[data-testid="stSidebar"] {
        padding-top: 30px;
    }
    /* Semua tombol menu */
    section[data-testid="stSidebar"] .stButton > button {
        width: 100%;
        height: 55px;
        border-radius: 8px;
        font-size: 24px;
        font-weight: 600;
        text-align: left;
        padding-left: 20px;
        margin-bottom: 8px;
    }
    /* -------------------------------------------------------
       JUDUL UTAMA
       ------------------------------------------------------- */
    .judul-dashboard {
        text-align: center;
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 30px;
    }
    /* -------------------------------------------------------
       JUDUL SECTION
       ------------------------------------------------------- */
    .judul-section {
        font-size: 22px;
        font-weight: 700;
        margin-bottom: 15px;
    }
    /* -------------------------------------------------------
       KARTU NILAI STOK
       ------------------------------------------------------- */
    .kartu-stok {
        border: 1px solid #d0d0d0;
        border-radius: 8px;
        padding: 18px;
        margin-bottom: 15px;
        text-align: center;
    }
    .judul-kartu {
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 10px;
    }
    .nilai-kartu {
        font-size: 28px;
        font-weight: 700;
    }
    .satuan-kartu {
        font-size: 16px;
        font-weight: 500;
        margin-top: 5px;
    }
</style>
""", unsafe_allow_html=True)

# CSS TABEL NERACA
st.markdown("""
<style>
    .neraca-container {
        width: 100%;
        overflow-x: auto;
        overflow-y: auto;
        max-height: 650px;
        border: 1px solid #d0d0d0;
    }
    /* Tabel */
    .neraca-table {
        border-collapse: collapse;
        width: max-content;
        min-width: 100%;
        font-size: 14px;
    }
    /* Semua cell */
    .neraca-table th,
    .neraca-table td {
        border: 1px solid #bdbdbd;
        padding: 8px 12px;
        white-space: nowrap;
        color: #222222;
    }
    /* Header WPP */
    .neraca-table .header-wpp {
        text-align: center;
        font-weight: 700;
        background-color: #eeeeee;
    }
    /* Header jenis ikan */
    .neraca-table .header-ikan {
        text-align: center;
        font-weight: 700;
        background-color: #f5f5f5;
    }
    /* Header komponen */
    .neraca-table .header-komponen {
        text-align: center;
        font-weight: 700;
        background-color: #eeeeee;
        position: sticky;
        left: 0;
        z-index: 4;
    }
    /* Kolom komponen */
    .neraca-table td:first-child {
        position: sticky;
        left: 0;
        background-color: white;
        z-index: 2;
    }
    /* Baris utama */
    .neraca-table .komponen-utama {
        font-weight: 700;
    }
    /* Subkomponen */
    .neraca-table .komponen-anak {
        font-weight: 600;
    }
    /* Sub-subkomponen */
    .neraca-table .komponen-subanak {
        font-weight: 400;
    }
    /* Nilai */
    .neraca-table .nilai-neraca {
        text-align: right;
        color: #222222;
        background-color: white;
    }
    /* Baris header tetap terlihat */
    .neraca-table thead th {
        position: sticky;
        top: 0;
        z-index: 3;
    }
    /* Header komponen harus paling depan */
    .neraca-table thead .header-komponen {
        z-index: 5;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 3. MEMBACA DATA
# ============================================================
# DATA STOK
file_data = r"Data\stok_dashboard.xlsx"
df = pd.read_excel(file_data)
# Membersihkan nama kolom
df.columns = df.columns.str.strip()
# Memastikan Tahun numerik
df["Tahun"] = pd.to_numeric(df["Tahun"], errors="coerce")
# Menghapus baris yang tidak memiliki Tahun atau WPP
df = df.dropna(subset=["Tahun", "WPP"])
# Mengubah Tahun menjadi integer
df["Tahun"] = df["Tahun"].astype(int)

# DATA TANGKAPAN
file_tangkap = r"Data\tangkapan_dashboard.xlsx"
df_tangkap = pd.read_excel(file_tangkap)
# Membersihkan nama kolom
df_tangkap.columns = df_tangkap.columns.str.strip()
# Memastikan Tahun numerik
df_tangkap["Tahun"] = pd.to_numeric(df_tangkap["Tahun"], errors="coerce")
# Menghapus baris yang tidak memiliki Tahun, WPP, dan Jenis Tangkapan
df_tangkap = df_tangkap.dropna(subset=["Tahun","WPP","jenis_tangkapan"])
# Mengubah Tahun menjadi integer
df_tangkap["Tahun"] = df_tangkap["Tahun"].astype(int)

# DATA NERACA
file_neraca = r"Data\neraca_dashboard.xlsx"
df_neraca = pd.read_excel(file_neraca)
# Membersihkan nama kolom
df_neraca.columns = df_neraca.columns.str.strip()
# Memastikan Tahun Numerik
df_neraca["Tahun"] = pd.to_numeric(df_neraca["Tahun"], errors="coerce")
# Menghapus baris yang tidak memiliki Tahun, WPP, dan Komponen
df_neraca = df_neraca.dropna(subset=["Tahun","Wilayah","Komponen"])
# Mengubah Tahun menjadi integer
df_neraca["Tahun"] = df_neraca["Tahun"].astype(int)


# ============================================================
# 4. DAFTAR JENIS IKAN
# ============================================================
# STOK
# Kolom pertama adalah Tahun
# Kolom kedua adalah WPP
# Sisanya dianggap sebagai jenis ikan
kolom_ikan = [
    kolom for kolom in df.columns
    if kolom not in ["Tahun", "WPP"]
]
# Nama yang ditampilkan pada dashboard
nama_ikan = {
    kolom: kolom.upper()
    for kolom in kolom_ikan
}

# TANGKAPAN
kolom_ikan_tangkap = [
    kolom_tangkap for kolom_tangkap in df_tangkap.columns
    if kolom_tangkap not in ["Tahun", "WPP", "jenis_tangkapan"]
]
# Nama yang ditampilkan pada dashboard
nama_ikan_tangkap = {
    kolom_tangkap: kolom_tangkap.upper()
    for kolom_tangkap in kolom_ikan_tangkap
}

# NERACA
kolom_ikan_neraca = [
    kolom_neraca for kolom_neraca in df_neraca.columns
    if kolom_neraca not in ["Tahun", "Wilayah", "Komponen"]
]
# Nama yang ditampilkan pada dashboard
nama_ikan_neraca = {
    kolom_neraca: kolom_neraca.upper()
    for kolom_neraca in kolom_ikan_neraca
}

# ============================================================
# 5. SIDEBAR / MENU
# ============================================================
if "halaman" not in st.session_state:
    st.session_state.halaman = "STOK"
# Fungsi untuk mengganti halaman
def pilih_halaman(nama_halaman):
    st.session_state.halaman = nama_halaman
with st.sidebar:
    # STOK
    st.button(
        "STOK",
        key="menu_stok",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.halaman == "STOK"
            else "secondary"
        ),
        on_click=pilih_halaman,
        args=("STOK",)
    )
    # TANGKAPAN
    st.button(
        "TANGKAPAN",
        key="menu_tangkapan",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.halaman == "TANGKAPAN"
            else "secondary"
        ),
        on_click=pilih_halaman,
        args=("TANGKAPAN",)
    )
    # NERACA
    st.button(
        "NERACA",
        key="menu_neraca",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.halaman == "NERACA"
            else "secondary"
        ),
        on_click=pilih_halaman,
        args=("NERACA",)
    )

# ============================================================
# 6. HALAMAN STOK
# ============================================================
if st.session_state.halaman == "STOK":
    # --------------------------------------------------------
    # JUDUL
    # --------------------------------------------------------
    st.markdown(
        '<div class="judul-dashboard">'
        'STOK IKAN DI PERAIRAN INDONESIA'
        '</div>',
        unsafe_allow_html=True
    )
    # --------------------------------------------------------
    # FILTER
    # --------------------------------------------------------
    col_tahun, col_wilayah, col_ikan = st.columns(
        [1, 1, 1]
    )
    # -------------------------
    # TAHUN
    # -------------------------
    with col_tahun:

        daftar_tahun = sorted(
            df["Tahun"].unique()
        )
        tahun_pilihan = st.multiselect(
            "TAHUN",
            options=daftar_tahun,
            default=daftar_tahun
        )
    # -------------------------
    # WILAYAH
    # -------------------------
    with col_wilayah:
        daftar_wpp = df["WPP"].unique().tolist()
        # Jika WPP-571 tersedia,
        # jadikan sebagai default
        if "WPP-571" in daftar_wpp:
            index_wpp = daftar_wpp.index("WPP-571")
        else:
            index_wpp = 0
        wilayah_pilihan = st.selectbox(
            "WILAYAH",
            options=daftar_wpp,
            index=index_wpp
        )
    # -------------------------
    # JENIS IKAN
    # -------------------------
    with col_ikan:
        jenis_ikan_pilihan = st.multiselect(
            "JENIS IKAN",
            options=kolom_ikan,
            default=kolom_ikan,
            format_func=lambda x: nama_ikan[x]
        )
     # ========================================================
    # FILTER DATA
    # ========================================================
    data_filter = df[
        (df["Tahun"].isin(tahun_pilihan)) &
        (df["WPP"] == wilayah_pilihan)
    ].copy()
    # ========================================================
    # MEMBUAT DATA LONG FORMAT
    # ========================================================
    if jenis_ikan_pilihan:
        data_long = data_filter.melt(
            id_vars=["Tahun", "WPP"],
            value_vars=jenis_ikan_pilihan,
            var_name="Jenis Ikan",
            value_name="Stok"
        )
        # Pastikan stok numerik
        data_long["Stok"] = pd.to_numeric(
            data_long["Stok"],
            errors="coerce"
        )
        # Hapus nilai kosong
        data_long = data_long.dropna(
            subset=["Stok"]
        )
    else:
        data_long = pd.DataFrame()
    # ========================================================
    # GRAFIK + RINGKASAN
    # ========================================================
    col_grafik, col_ringkasan = st.columns(
        [2.2, 1]
    )
    # ========================================================
    # GRAFIK PERKEMBANGAN STOK
    # ========================================================
    with col_grafik:
        st.markdown(
            '<div class="judul-section">'
            'PERKEMBANGAN STOK'
            '</div>',
            unsafe_allow_html=True
        )
        if not data_long.empty:
            fig = px.line(
                data_long,
                x="Tahun",
                y="Stok",
                color="Jenis Ikan",
                markers=True,
                labels={
                    "Tahun": "TAHUN",
                    "Stok": "VOLUME STOK (ton)",
                    "Jenis Ikan": "JENIS IKAN"
                }
            )
            fig.update_layout(
                height=500,
                margin=dict(
                    l=20,
                    r=20,
                    t=20,
                    b=20
                ),
                legend_title_text="Jenis Ikan",
                xaxis=dict(
                    dtick=1
                )
            )
            st.plotly_chart(
                fig,
                use_container_width=True
            )
        else:
            st.info(
                "Pilih minimal satu TAHUN dan satu JENIS IKAN."
            )
     # ========================================================
    # NILAI STOK
    # ========================================================
    with col_ringkasan:
        st.markdown(
            '<div class="judul-section">'
            'VOLUME STOK'
            '</div>',
            unsafe_allow_html=True
        )
        if not data_long.empty:
            # -----------------------------------------------
            # TOTAL
            # -----------------------------------------------
            total_stok = data_long["Stok"].sum()
            st.markdown(
                f"""
                <div class="kartu-stok">
                    <div class="judul-kartu">
                        TOTAL
                    </div>
                    <div class="nilai-kartu">
                        {total_stok:,.2f}
                    </div>
                    <div class="satuan-kartu">
                        ton
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
            # -----------------------------------------------
            # RATA-RATA
            # -----------------------------------------------
            rata_stok = data_long["Stok"].mean()
            st.markdown(
                f"""
                <div class="kartu-stok">
                    <div class="judul-kartu">
                        RATA-RATA
                    </div>
                    <div class="nilai-kartu">
                        {rata_stok:,.2f}
                    </div>
                    <div class="satuan-kartu">
                        ton
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info(
                "Belum ada data stok."
            )

# ============================================================
# 7. HALAMAN TANGKAPAN
# ============================================================
elif st.session_state.halaman == "TANGKAPAN":
    # --------------------------------------------------------
    # JUDUL
    # --------------------------------------------------------
    st.markdown(
        '<div class="judul-dashboard">'
        'TANGKAPAN IKAN DI PERAIRAN INDONESIA'
        '</div>',
        unsafe_allow_html=True
    )
    # --------------------------------------------------------
    # FILTER
    # --------------------------------------------------------
    col_tahun_tangkap, col_wilayah_tangkap, col_tangkap = st.columns(
        [1, 1, 1]
    )
    # -------------------------
    # TAHUN
    # -------------------------
    with col_tahun_tangkap:

        daftar_tahun_tangkap = sorted(
            df_tangkap["Tahun"].unique()
        )
        tahun_pilihan_tangkap = st.multiselect(
            "TAHUN",
            options=daftar_tahun_tangkap,
            default=daftar_tahun_tangkap
        )
    # -------------------------
    # WILAYAH
    # -------------------------
    with col_wilayah_tangkap:
        daftar_wpp_tangkap = df_tangkap["WPP"].unique().tolist()
        # Jika WPP-571 tersedia,
        # jadikan sebagai default
        if "WPP-571" in daftar_wpp_tangkap:
            index_wpp_tangkap = daftar_wpp_tangkap.index("WPP-571")
        else:
            index_wpp_tangkap = 0
        wilayah_pilihan_tangkap = st.selectbox(
            "WILAYAH",
            options=daftar_wpp_tangkap,
            index=index_wpp_tangkap
        )
    # -------------------------
    # JENIS TANGKAPAN
    # -------------------------
    with col_tangkap:
        daftar_jenis_tangkapan = (
            df_tangkap["jenis_tangkapan"]
            .dropna()
            .unique()
            .tolist()
        )


        jenis_tangkapan_pilihan = st.multiselect(
            "JENIS TANGKAPAN",
            options=daftar_jenis_tangkapan,
            default=daftar_jenis_tangkapan,
            format_func=lambda x: x.upper()
        )
    # ========================================================
    # FILTER DATA TANGKAPAN
    # ========================================================

    data_tangkapan_filter = df_tangkap[
        (
            df_tangkap["Tahun"].isin(tahun_pilihan_tangkap)
        )
        &
        (
            df_tangkap["WPP"] == wilayah_pilihan_tangkap
        )
        &
        (
            df_tangkap["jenis_tangkapan"].isin(jenis_tangkapan_pilihan)
        )
    ].copy()

    # ========================================================
    # MENGUBAH KOLOM IKAN MENJADI NUMERIK
    # ========================================================
    for kolom_tangkap in kolom_ikan_tangkap:
        data_tangkapan_filter[kolom_tangkap] = pd.to_numeric(
            data_tangkapan_filter[kolom_tangkap],
            errors="coerce"
        )

    # ========================================================
    # TOTAL TANGKAPAN PER BARIS
    # ========================================================
    if not data_tangkapan_filter.empty:
        data_tangkapan_filter[
            "Total Tangkapan"
        ] = (
            data_tangkapan_filter[
                kolom_ikan_tangkap
            ]
            .sum(axis=1)
        )
    else:
        data_tangkapan_filter[
            "Total Tangkapan"
        ] = 0

    # ========================================================
    # GRAFIK + RINGKASAN
    # ========================================================
    col_grafik_tangkapan, col_ringkasan_tangkapan = (
        st.columns([2.2, 1])
    )
    # ========================================================
    # GRAFIK PERKEMBANGAN TANGKAPAN
    # ========================================================
    with col_grafik_tangkapan:
        st.markdown(
            '<div class="judul-section">'
            'PERKEMBANGAN TANGKAPAN'
            '</div>',
            unsafe_allow_html=True
        )
        if not data_tangkapan_filter.empty:
            fig_tangkapan = px.bar(
                data_tangkapan_filter,
                x="Tahun",
                y="Total Tangkapan",
                color="jenis_tangkapan",
                barmode="group",
                labels={
                    "Tahun": "TAHUN",
                    "Total Tangkapan":
                        "VOLUME TANGKAPAN",
                    "jenis_tangkapan":
                        "JENIS TANGKAPAN"
                }
            )
            fig_tangkapan.update_layout(
                height=500,
                margin=dict(
                    l=20,
                    r=20,
                    t=20,
                    b=20
                ),
                legend_title_text=
                    "Jenis Tangkapan",
                xaxis=dict(
                    dtick=1
                )
            )
            st.plotly_chart(
                fig_tangkapan,
                use_container_width=True
            )
        else:
            st.info(
                "Pilih minimal satu TAHUN atau satu JENIS TANGKAPAN."
            )
    # ========================================================
    # NILAI TANGKAPAN
    # ========================================================
    with col_ringkasan_tangkapan:
        st.markdown(
            '<div class="judul-section">'
            'VOLUME TANGKAPAN'
            '</div>',
            unsafe_allow_html=True
        )
        if not data_tangkapan_filter.empty:
            # ------------------------------------------------
            # TOTAL
            # ------------------------------------------------
            total_tangkapan = (
                data_tangkapan_filter[
                    "Total Tangkapan"
                ].sum()
            )
            st.markdown(
                f"""
                <div class="kartu-stok">
                    <div class="judul-kartu">
                        TOTAL
                    </div>
                    <div class="nilai-kartu">
                        {total_tangkapan:,.2f}
                    </div>
                    <div class="satuan-kartu">
                        ton
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
            # ------------------------------------------------
            # RATA-RATA
            # ------------------------------------------------
            rata_tangkapan = (
                data_tangkapan_filter[
                    "Total Tangkapan"
                ].mean()
            )
            st.markdown(
                f"""
                <div class="kartu-stok">
                    <div class="judul-kartu">
                        RATA-RATA
                    </div>
                    <div class="nilai-kartu">
                        {rata_tangkapan:,.2f}
                    </div>
                    <div class="satuan-kartu">
                        ton
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info(
                "Belum ada data tangkapan."
            )

# ============================================================
# 8. HALAMAN NERACA
# ============================================================
elif st.session_state.halaman == "NERACA":
    st.markdown(
        '<div class="judul-dashboard">'
        'NERACA ASET SUMBER DAYA AKUATIK'
        '</div>',
        unsafe_allow_html=True
    )
    # ========================================================
    # FILTER
    # ========================================================
    col_tahun_neraca, col_wilayah_neraca, col_ikan_neraca = st.columns(
        [1, 1, 1]
    )
    # ========================================================
    # TAHUN
    # ========================================================
    with col_tahun_neraca:
        daftar_tahun_neraca = sorted(df_neraca["Tahun"].unique())
        if 2020 in daftar_tahun_neraca:
            index_tahun_neraca = daftar_tahun_neraca.index(2020)
        else:
            index_tahun_neraca = 0
        tahun_neraca_pilihan = st.selectbox(
            "TAHUN",
            options=daftar_tahun_neraca,
            index=index_tahun_neraca
        )
    # ========================================================
    # WILAYAH
    # ========================================================
    with col_wilayah_neraca:
        daftar_wpp_neraca = (
            df_neraca["Wilayah"]
            .dropna()
            .unique()
            .tolist()
        )
        wilayah_neraca_pilihan = st.multiselect(
            "WILAYAH",
            options=daftar_wpp_neraca,
            default=daftar_wpp_neraca
        )
    # ========================================================
    # JENIS IKAN
    # ========================================================
    with col_ikan_neraca:
        jenis_ikan_neraca_pilihan = st.multiselect(
            "JENIS IKAN",
            options=kolom_ikan_neraca,
            default=kolom_ikan_neraca,
            format_func=lambda x: nama_ikan_neraca[x]
        )
    # ========================================================
    # FILTER DATA
    # ========================================================
    data_neraca_filter = df_neraca[
        (
            df_neraca["Tahun"] == tahun_neraca_pilihan
        )
        &
        (
            df_neraca["Wilayah"].isin(wilayah_neraca_pilihan)
        )
    ].copy()
    # ========================================================
    # URUTAN KOMPONEN NERACA
    # ========================================================
    urutan_komponen = [
        "STOK AWAL TAHUN",
        "PENAMBAHAN STOK",
        "PENGURANGAN STOK",
        "Tangkapan",
        "Industri",
        "Non-Industri",
        "Kematian Alami Ikan",
        "STOK AKHIR TAHUN"
    ]
    # ========================================================
    # MENENTUKAN KOMPONEN YANG ADA
    # ========================================================
    komponen_tersedia = [
        komponen
        for komponen in urutan_komponen
        if komponen in (
            data_neraca_filter["Komponen"]
            .astype(str)
            .tolist()
        )
    ]
    # ========================================================
    # FUNGSI UNTUK MENENTUKAN LEVEL KOMPONEN
    # ========================================================
    def level_komponen(komponen):
        if komponen in [
            "STOK AWAL TAHUN",
            "PENAMBAHAN STOK",
            "PENGURANGAN STOK",
            "STOK AKHIR TAHUN"
        ]:
            return 0
        elif komponen in [
            "Tangkapan",
            "Kematian Alami Ikan"
        ]:
            return 1
        elif komponen in [
            "Industri",
            "Non-Industri"
        ]:
            return 2
        else:
            return 0
    # ========================================================
    # TABEL NERACA
    # ========================================================
    st.markdown(
        '<div class="judul-section">'
        'dalam satuan ton'
        '</div>',
        unsafe_allow_html=True
    )
    # ========================================================
    # PERSIAPAN TABEL DOWNLOAD
    # ========================================================
    def buat_excel_neraca(
        data_neraca,
        tahun,
        wilayah_pilihan,
        ikan_pilihan,
        nama_ikan
    ):
        wb = Workbook()
        ws = wb.active
        ws.title = "Neraca"
        # ========================================================
        # STYLE
        # ========================================================
        thin = Side(
            style="thin",
            color="B7B7B7"
        )
        border = Border(
            left=thin,
            right=thin,
           top=thin,
            bottom=thin
        )
        font_normal = Font(
            name="Arial",
            size=10
        )
        font_bold = Font(
            name="Arial",
            size=10,
            bold=True
        )
        center = Alignment(
            horizontal="center",
            vertical="center"
        )
        right = Alignment(
            horizontal="right",
            vertical="center"
        )
        left = Alignment(
            horizontal="left",
            vertical="center"
        )
        # ========================================================
        # JUDUL
        # ========================================================
        jumlah_kolom = (
            1 + len(wilayah_pilihan) * len(ikan_pilihan)
        )
        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=jumlah_kolom
        )
        ws.cell(
            row=1,
            column=1,
            value=(
                f"NERACA ASET SUMBER DAYA AKUATIK (dalam satuan ton) - "
                f"{tahun}"
            )
        )
        ws.cell(
            row=1,
            column=1
        ).font = Font(
            name="Arial",
            size=14,
            bold=True
        )
        ws.cell(
            row=1,
            column=1
        ).alignment = center
        # ========================================================
        # HEADER KOMPONEN
        # ========================================================
        ws.merge_cells(
            start_row=3,
            start_column=1,
            end_row=4,
            end_column=1
        )
        cell = ws.cell(
            row=3,
            column=1,
            value="KOMPONEN"
        )
        cell.font = font_bold
        cell.alignment = center
        cell.border = border
        # ========================================================
        # HEADER WPP DAN JENIS IKAN
        # ========================================================
        kolom_excel = 2
        for wpp in wilayah_pilihan:
            jumlah_ikan = len(ikan_pilihan)
            # Merge WPP
            ws.merge_cells(
                start_row=3,
                start_column=kolom_excel,
                end_row=3,
                end_column=(
                    kolom_excel + jumlah_ikan - 1
                )
            )
            cell_wpp = ws.cell(
                row=3,
                column=kolom_excel,
                value=wpp
            )
            cell_wpp.font = font_bold
            cell_wpp.alignment = center
            # Border semua cell header WPP
            for c in range(
                kolom_excel,
                kolom_excel + jumlah_ikan
            ):
                ws.cell(
                    row=3,
                    column=c
                ).border = border
            # Jenis ikan
            for ikan in ikan_pilihan:
                cell_ikan = ws.cell(
                    row=4,
                    column=kolom_excel,
                    value=nama_ikan[ikan]
                )
                cell_ikan.font = font_bold
                cell_ikan.alignment = center
                cell_ikan.border = border
                kolom_excel += 1
        # ========================================================
        # KOMPONEN
        # ========================================================
        urutan_komponen = [
            "STOK AWAL TAHUN",
            "PENAMBAHAN STOK",
            "PENGURANGAN STOK",
            "Tangkapan",
            "Industri",
            "Non-Industri",
            "Kematian Alami Ikan",
            "STOK AKHIR TAHUN"
        ]
        def level_komponen(komponen):
            if komponen in [
                "STOK AWAL TAHUN",
                "PENAMBAHAN STOK",
                "PENGURANGAN STOK",
                "STOK AKHIR TAHUN"
            ]:
                return 0
            elif komponen in [
                "Tangkapan",
                "Kematian Alami Ikan"
            ]:
                return 1
            elif komponen in [
                "Industri",
                "Non-Industri"
            ]:
                return 2
            return 0
        # ========================================================
        # ISI DATA
        # ========================================================
        baris_excel = 5
        for komponen in urutan_komponen:
            data_komponen = data_neraca[
                data_neraca["Komponen"]
                == komponen
            ]
            if data_komponen.empty:
                continue
            # ----------------------------------------------------
            # Nama komponen
            # ----------------------------------------------------
            level = level_komponen(
                komponen
            )
            cell = ws.cell(
                row=baris_excel,
                column=1,
                value=komponen
            )
            cell.font = font_bold if level == 0 else font_normal
            cell.alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=level
            )
            cell.border = border
            # ----------------------------------------------------
            # Nilai setiap WPP × jenis ikan
            # ----------------------------------------------------
            kolom_excel = 2
            for wpp in wilayah_pilihan:
                data_wpp = data_komponen[
                    data_komponen["Wilayah"]
                    == wpp
                ]
                for ikan in ikan_pilihan:
                    if (
                        not data_wpp.empty
                        and ikan in data_wpp.columns
                    ):
                        nilai = pd.to_numeric(
                            data_wpp.iloc[0][ikan],
                            errors="coerce"
                        )
                    else:
                        nilai = None
                    cell_nilai = ws.cell(
                        row=baris_excel,
                        column=kolom_excel,
                        value=(
                            float(nilai)
                            if pd.notna(nilai)
                            else None
                        )
                    )
                    cell_nilai.number_format = (
                        '#,##0.00'
                    )
                    cell_nilai.alignment = right
                    cell_nilai.border = border
                    kolom_excel += 1
            baris_excel += 1
        # ========================================================
        # LEBAR KOLOM
        # ========================================================
        ws.column_dimensions["A"].width = 25
        for col in range(
            2,
            jumlah_kolom + 1
        ):
            letter = get_column_letter(col)
            ws.column_dimensions[
                letter
            ].width = 18
        # ========================================================
        # FREEZE PANES
        # ========================================================
        ws.freeze_panes = "B5"
        # ========================================================
        # TINGGI BARIS
        # ========================================================
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[3].height = 22
        ws.row_dimensions[4].height = 22
        # ========================================================
        # SIMPAN KE MEMORY
        # ========================================================
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output     

    # ========================================================
    # DOWNLOAD BUTTON
    # ========================================================
    col_kosong, col_download = st.columns(
        [5, 1]
    )
    with col_download:
        # ----------------------------------------------------
        # Data yang akan didownload
        # ----------------------------------------------------    
        file_excel = buat_excel_neraca(
            data_neraca=data_neraca_filter,
            tahun=tahun_neraca_pilihan,
            wilayah_pilihan=wilayah_neraca_pilihan,
            ikan_pilihan=jenis_ikan_neraca_pilihan,
            nama_ikan=nama_ikan_neraca
        )
        st.download_button(
            label="DOWNLOAD",
            data=file_excel,
            file_name=(
                f"Neraca Aset Sumber Daya Akuatik_{tahun_neraca_pilihan}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True
        )
        
    # ========================================================
    # MEMBUAT TABEL HTML
    # ========================================================
    if (
        not data_neraca_filter.empty
        and wilayah_neraca_pilihan
        and jenis_ikan_neraca_pilihan
    ):
        # ----------------------------------------------------
        # Mulai tabel
        # ----------------------------------------------------
        tabel_html = """
        <div class="neraca-container">
        <table class="neraca-table">
        """
        # ====================================================
        # HEADER BARIS 1
        # TAHUN + WPP
        # ====================================================
        tabel_html += """
        <thead>
        <tr>
            <th
                class="header-komponen"
                rowspan="2"
            >
                KOMPONEN
            </th>
        """
        # Setiap WPP menjadi satu kelompok
        for wpp in wilayah_neraca_pilihan:
            tabel_html += f"""
            <th
                class="header-wpp"
                colspan="{len(jenis_ikan_neraca_pilihan)}"
            >
                {html.escape(str(wpp))}
            </th>
            """
        tabel_html += """
        </tr>
        <tr>
        """
        # ====================================================
        # HEADER BARIS 2
        # JENIS IKAN
        # ====================================================
        for wpp in wilayah_neraca_pilihan:
            for ikan in jenis_ikan_neraca_pilihan:
                tabel_html += f"""
                <th class="header-ikan">
                    {html.escape(
                        nama_ikan_neraca[ikan]
                    )}
                </th>
                """
        tabel_html += """
        </tr>
        </thead>
        <tbody>
        """
        # ====================================================
        # ISI TABEL
        # ====================================================
        for komponen in komponen_tersedia:
            level = level_komponen(
                komponen
            )
            # ------------------------------------------------
            # Style berdasarkan level
            # ------------------------------------------------
            if level == 0:
                kelas_komponen = (
                    "komponen-utama"
                )
            elif level == 1:
                kelas_komponen = (
                    "komponen-anak"
                )
            else:
                kelas_komponen = (
                    "komponen-subanak"
                )
            # ------------------------------------------------
            # Nama komponen
            # ------------------------------------------------
            if level == 0:
                nama_tampilan = html.escape(
                    komponen
                )
            elif level == 1:
                nama_tampilan = (
                    "&nbsp;&nbsp;&nbsp;"
                    + html.escape(komponen)
                )
            else:
                nama_tampilan = (
                    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    + html.escape(komponen)
                )
            tabel_html += f"""
            <tr>
                <td class="{kelas_komponen}">
                    {nama_tampilan}
                </td>
            """
            # ------------------------------------------------
            # Nilai per WPP dan jenis ikan
            # ------------------------------------------------
            for wpp in wilayah_neraca_pilihan:
                data_wpp = data_neraca_filter[
                    data_neraca_filter[
                        "Wilayah"
                    ] == wpp
                ]
                for ikan in jenis_ikan_neraca_pilihan:
                    data_nilai = data_wpp[
                        data_wpp["Komponen"]
                        == komponen
                    ]
                    if (
                        not data_nilai.empty
                        and ikan in data_nilai.columns
                    ):
                        nilai = pd.to_numeric(
                            data_nilai.iloc[0][ikan],
                            errors="coerce"
                        )
                    else:
                        nilai = None
                    if pd.isna(nilai):
                        nilai_tampil = "-"
                    else:
                        nilai_tampil = (
                            f"{nilai:,.2f}"
                        )
                    tabel_html += f"""
                    <td class="nilai-neraca">
                        {nilai_tampil}
                    </td>
                    """
            tabel_html += """
            </tr>
            """
        # ====================================================
        # SELESAI TABEL
        # ====================================================
        tabel_html += """
        </tbody>
        </table>
        </div>
        """
        # Menampilkan tabel
        st.html(tabel_html)
    elif not wilayah_neraca_pilihan:
        st.info(
            "Pilih minimal satu WILAYAH."
        )
    elif not jenis_ikan_neraca_pilihan:
        st.info(
            "Pilih minimal satu JENIS IKAN."
        )
    else:
        st.warning(
            "Data neraca tidak ditemukan "
            "untuk pilihan tersebut."
        )