"""Streamlit interface for processing Accurate delivery-note workbooks."""

from datetime import datetime

import pandas as pd # pyright: ignore[reportMissingModuleSource]
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import streamlit as st
from PIL import Image, UnidentifiedImageError

from config.perusahaan import KONFIGURASI_PERUSAHAAN
from processors.cleaner import clean_data
from processors.exporter import export_to_excel
from processors.filters import filter_zero_prices
from processors.reader import read_excel_file
from processors.summaries import (
    generate_rekap_barang,
    generate_rekap_pelanggan,
)
from processors.volume import REVIEW_STATUS, REVIEW_STATUS_COLUMN, calculate_bata_m3


st.set_page_config(page_title="Otomatisasi Surat Jalan", page_icon="📄", layout="wide")

# === SISTEM NAVIGASI WELCOME PAGE ===
if "app_state" not in st.session_state:
    st.session_state["app_state"] = "welcome"

if st.session_state["app_state"] == "welcome":
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    .stApp { 
        background: linear-gradient(135deg, #F8FAFC 0%, #F1F5F9 100%);
        font-family: 'Inter', -apple-system, sans-serif;
    }
    header {visibility: hidden;}
    
    .hero-wrapper {
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        margin-top: 25vh; 
        margin-bottom: 35px; 
        position: relative;
    }

    .hero-content { z-index: 10; position: relative; text-align: center; }

    .hero-title-small { 
        font-size: 22px; font-weight: 500; color: #64748B; 
        margin-bottom: -5px; letter-spacing: 0.5px;
    }
    
    .hero-title-large {
        font-size: 68px; font-weight: 800; letter-spacing: -2.5px; line-height: 1.2; 
        margin-bottom: 0px; color: #0F172A; 
    }
    .text-highlight { color: #4338CA; } 
    
    /* 1. Kunci Utama Penengah (Super Stabil) */
    div[data-testid="stButton"] { 
        text-align: center !important;
        display: block !important;
        width: 100% !important;
    }
    div[data-testid="stElementContainer"]:has(> div[data-testid="stButton"]) {
        width: 100% !important;
    }
    
    /* 2. Styling Tombol Presisi (Bentuk Kapsul Terjaga) */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #4F46E5 0%, #4338CA 100%) !important;
        border-radius: 50px !important; 
        border: none !important; 
        min-width: 250px !important; 
        padding: 16px 40px !important; 
        box-shadow: 0 8px 25px -4px rgba(67, 56, 202, 0.4) !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        
        /* Trik agar terkunci di tengah */
        display: inline-flex !important; 
        justify-content: center !important;
        align-items: center !important;
        width: auto !important;
        flex: 0 0 auto !important;
        margin: 0 auto !important;
    }
    
    /* 3. Teks Font Tombol Besar & Presisi */
    div[data-testid="stButton"] > button p {
        font-weight: 700 !important; 
        font-size: 22px !important; 
        letter-spacing: 1.5px !important;
        color: #ffffff !important;
        margin: 0 !important;
    }
    
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-4px) !important; 
        box-shadow: 0 12px 30px -4px rgba(67, 56, 202, 0.55) !important;
        background: linear-gradient(135deg, #4338CA 0%, #3730A3 100%) !important;
    }
    div[data-testid="stButton"] > button:active {
        transform: scale(0.97) !important;
    }
    </style>

    <div class="hero-wrapper">
        <div class="hero-content">
            <div class="hero-title-small">Selamat Datang di</div>
            <div class="hero-title-large">ClearFlow <span class="text-highlight">System</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Tombol dipanggil langsung tanpa pembungkus kolom
    if st.button("Mulai"):
        st.session_state["app_state"] = "main"
        st.rerun()
            
    st.stop()

# === MANAJEMEN STATE PERUSAHAAN ===
if "selected_company" not in st.session_state:
    st.session_state["selected_company"] = KONFIGURASI_PERUSAHAAN["nama"]

# === ADVANCED CSS INJECTION ===
st.markdown("""
<style>
/* Font & Latar Belakang */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp { 
    background-color: #F8FAFC; 
    font-family: 'Inter', -apple-system, sans-serif;
}

/* Floating Top Bar (Glassmorphism) */
.top-bar {
    background: rgba(255, 255, 255, 0.85); 
    padding: 22px 35px; 
    border-radius: 16px;
    border: 1px solid rgba(255, 255, 255, 0.4);
    display: flex; justify-content: space-between; align-items: center; 
    margin-top: 10px; margin-bottom: 45px;
    box-shadow: 0 4px 25px -2px rgba(0, 0, 0, 0.04);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
}
.brand-text {
    font-weight: 800; font-size: 25px; letter-spacing: -0.5px;
    background: linear-gradient(135deg, #4338CA 0%, #7C3AED 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.brand-ht { font-size: 13px; font-weight: 700; color: #E11D48; margin-right: 4px; vertical-align: top; }

/* 1. BAGIAN PUTIH KARTU (HTML) */
.card-ui {
    background-color: #ffffff; 
    border: 1px solid #F1F5F9; 
    border-bottom: none; 
    border-radius: 16px 16px 0 0;
    padding: 40px 10px 30px 10px; 
    display: flex; flex-direction: column; align-items: center;
    box-shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.03);
    transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
}
.logo-circle {
    width: 82px; height: 82px; border-radius: 50%; margin-bottom: 25px;
    display: flex; justify-content: center; align-items: center; font-size: 32px; color: white;
    box-shadow: 0 8px 20px -4px rgba(0,0,0,0.15), 0 0 0 5px rgba(255,255,255,0.8);
    border: 1px solid rgba(0,0,0,0.05);
}
.bg-sangnila { background-color: #1E293B; }
.bg-mpi { background-color: #4D7C0F; }
.bg-tsmj { background-color: #C2410C; }
.tag-commercial { 
    font-size: 10.5px; color: #4338CA; background: #EEF2FF; 
    padding: 6px 16px; border-radius: 20px; font-weight: 700; 
    letter-spacing: 0.8px; text-transform: uppercase;
    border: 1px solid #E0E7FF;
}

/* 2. TOMBOL UMUM & KARTU (Deep Indigo Premium) */
div[data-testid="stButton"] button {
    background: linear-gradient(135deg, #4F46E5 0%, #6366F1 100%) !important;
    color: #ffffff !important;
    border-radius: 10px !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    padding: 12px 28px !important;
    font-weight: 600 !important; font-size: 14px !important; letter-spacing: 0.3px !important;
    box-shadow: 0 4px 14px 0 rgba(79, 70, 229, 0.25) !important;
    transition: all 0.3s ease !important;
}
div[data-testid="stButton"] button:hover {
    background: linear-gradient(135deg, #4338CA 0%, #4F46E5 100%) !important;
    box-shadow: 0 6px 20px rgba(67, 56, 202, 0.4) !important;
    transform: translateY(-2px) !important;
}
div[data-testid="stButton"] button:active { transform: scale(0.97) !important; }

/* 3. SETTING KHUSUS TOMBOL BAWAH KARTU */
div[data-testid="column"] div[data-testid="stButton"] {
    margin-top: -1.2rem !important; 
    position: relative; z-index: 10;
}
div[data-testid="column"] div[data-testid="stButton"] button {
    border-radius: 0 0 16px 16px !important;
    border-top: 1px solid rgba(255,255,255,0.2) !important;
    padding: 16px 5px !important;
}
div[data-testid="column"] div[data-testid="stButton"] button:hover { transform: none !important; }

/* Efek pop-up */
div[data-testid="column"] { transition: transform 0.4s cubic-bezier(0.16, 1, 0.3, 1); }
div[data-testid="column"]:hover { transform: translateY(-8px); }
div[data-testid="column"]:hover .card-ui { box-shadow: 0 15px 35px -5px rgba(0,0,0,0.06); }

/* Notifikasi Elegan */
.active-alert {
    background: linear-gradient(to right, #ffffff, #F8FAFC);
    border: 1px solid #E2E8F0; border-left: 5px solid #4F46E5;
    padding: 18px 25px; border-radius: 12px; 
    color: #1E293B; font-weight: 500; font-size: 15px;
    margin-top: 5px; margin-bottom: 30px; 
    box-shadow: 0 4px 15px -3px rgba(0,0,0,0.02);
    display: flex; align-items: center;
}
.active-alert-highlight { font-weight: 800; color: #4338CA; margin-left: 6px; }
</style>
""", unsafe_allow_html=True)

# === TOP NAVBAR (Warna teks disesuaikan dengan tema ungu) ===
st.markdown("""
<div class="top-bar">
    <div class="brand-text"><span class="brand-ht">HT</span>ClearFlow System</div>
    <div style="color: #64748B; font-size: 13.5px; font-weight: 500;">Sales Marketing | <b style="color: #334155;">Otomatisasi Surat Jalan</b></div>
</div>
""", unsafe_allow_html=True)

# === COMPANY CARDS GRID ===
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class="card-ui">
        <div class="logo-circle bg-sangnila">&#9884;</div>
        <div class="tag-commercial">COMMERCIAL</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("CV. SANGNILA TIRTA UTAMA", use_container_width=True):
        st.session_state["selected_company"] = "CV Sangnila Tirta Utama"

with col2:
    st.markdown("""
    <div class="card-ui">
        <div class="logo-circle bg-mpi">&#10034;</div>
        <div class="tag-commercial">COMMERCIAL</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("PT. MAJU PERKASA INDONESIA", use_container_width=True):
        st.session_state["selected_company"] = "PT Maju Perkasa Indonesia"

with col3:
    st.markdown("""
    <div class="card-ui">
        <div class="logo-circle bg-tsmj">&#9735;</div>
        <div class="tag-commercial">COMMERCIAL</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("PT. TIRTA SARI MULIA JAYA", use_container_width=True):
        st.session_state["selected_company"] = "PT Tirta Sari Mulia Jaya"

selected_company = st.session_state["selected_company"]
st.markdown(f"""
<div class="active-alert">
    <span>Perusahaan yang aktif: </span>
    <span class="active-alert-highlight">{selected_company}</span>
</div>
""", unsafe_allow_html=True)
st.write("---")

uploaded_file = st.file_uploader(
    "Unggah file laporan Excel",
    type=("xlsx", "xlsm"),
    help="Format yang didukung: .xlsx dan .xlsm.",
)
process_clicked = st.button(
    "Proses Data", type="primary", disabled=uploaded_file is None
)

proses_perusahaan = False
if selected_company == "PT Maju Perkasa Indonesia":
    company_config = KONFIGURASI_PERUSAHAAN
    proses_perusahaan = True
elif selected_company == "CV Sangnila Tirta Utama":
    company_config = {
        "kolom_wajib_asli": [
            "Pelanggan",
            "Nama Barang",
            "UOM",
            "Kuantitas",
            "Rasio/Fraksi",
            "Total Harga",
            "@Harga",
        ],
        "kolom_wajib": [
            "Nama Pelanggan",
            "Barang & Jasa",
            "Satuan",
            "Kuantitas",
            "Rasio Satuan #2 Barang & Jasa",
            "Total Harga",
            "@Harga",
        ],
    }
    proses_perusahaan = True
elif selected_company == "PT Tirta Sari Mulia Jaya":
    company_config = {
        "kolom_wajib": [
            "Nama Pelanggan",
            "Barang & Jasa",
            "Satuan",
            "Kuantitas",
            "Total Harga",
            "@Harga",
        ],
    }
    proses_perusahaan = True

if proses_perusahaan and uploaded_file is not None:
    st.caption(
        f"File: {uploaded_file.name} · "
        f"{uploaded_file.size / 1024:.1f} KB"
    )

    try:
        baris_blank_sangnila = 0
        if selected_company == "PT Tirta Sari Mulia Jaya":
            df_raw = pd.read_excel(uploaded_file)
            df_raw.columns = df_raw.columns.astype(str).str.strip()
            df_raw.rename(
                columns={
                    "Pelanggan": "Nama Pelanggan",
                    "Nama Barang": "Barang & Jasa",
                    "UOM": "Satuan",
                },
                inplace=True,
            )
            preview_df = df_raw
        else:
            req_cols = company_config.get(
                "kolom_wajib_asli", company_config["kolom_wajib"]
            )
            if selected_company == "CV Sangnila Tirta Utama":
                preview_df, _ = read_excel_file(
                    uploaded_file, required_columns=req_cols
                )
            else:
                preview_df, _ = read_excel_file(uploaded_file)

        preview_df.columns = preview_df.columns.astype(str).str.strip()

        if selected_company == "CV Sangnila Tirta Utama":
            jumlah_baris_sebelum_filter = len(preview_df)
            preview_df = preview_df.rename(
                columns={
                    "Pelanggan": "Nama Pelanggan",
                    "Nama Barang": "Barang & Jasa",
                    "UOM": "Satuan",
                    "Rasio/Fraksi": "Rasio Satuan #2 Barang & Jasa",
                    "Alamat Pengiriman": "Alamat pengiriman",
                }
            )
            # Tambahkan kolom dummy agar fungsi calculate_bata_m3 milik PT MPI tidak error
            preview_df["Alamat pengiriman"] = ""

            # --- OVERRIDE RASIO AXEL AAC ---
            rasio_pengganti = {
                "AXEL AAC 10X20X40CM": 83,
                "AXEL AAC 10X20X50CM": 100,
                "AXEL AAC 10X20X58CM": 83,
                "AXEL AAC 10X20X60CM": 83,
                "AXEL AAC 7,5X20X40CM": 111,
                "AXEL AAC 7,5X20X50CM": 111,
                "AXEL AAC 7,5X20X60CM": 111,
            }

            # Jika Barang & Jasa ada di kamus rasio_pengganti, timpa nilai rasionya
            mask_axel = preview_df["Barang & Jasa"].isin(rasio_pengganti.keys())
            if mask_axel.any():
                preview_df.loc[
                    mask_axel, "Rasio Satuan #2 Barang & Jasa"
                ] = preview_df.loc[mask_axel, "Barang & Jasa"].map(rasio_pengganti)

            kolom_cek_blank = [
                "Total Harga",
                "@Harga",
                "Barang & Jasa",
                "Nama Pelanggan",
            ]
            kolom_tersedia = [
                col for col in kolom_cek_blank if col in preview_df.columns
            ]
            preview_df = preview_df.dropna(subset=kolom_tersedia)
            for col in kolom_tersedia:
                preview_df = preview_df[
                    preview_df[col].astype(str).str.strip() != ""
                ]
            baris_blank_sangnila = jumlah_baris_sebelum_filter - len(preview_df)
    except (ValueError, OSError, ImportError) as error:
        st.error(f"File tidak dapat dibaca atau strukturnya tidak valid: {error}")
    else:
        if process_clicked:
            try:
                initial_rows = len(preview_df) + baris_blank_sangnila
                cleaned_input = clean_data(preview_df)
                blank_rows_removed = initial_rows - len(cleaned_input)
                for column in company_config["kolom_wajib"]:
                    if column not in cleaned_input.columns:
                        cleaned_input[column] = pd.NA
                cleaned_input = cleaned_input.loc[
                    :,
                    [
                        column
                        for column in preview_df.columns
                        if column in cleaned_input.columns
                    ],
                ]
                if selected_company == "PT Tirta Sari Mulia Jaya":
                    calculated_df = cleaned_input.copy()
                    calculated_df[REVIEW_STATUS_COLUMN] = "LULUS"
                else:
                    calculated_df = calculate_bata_m3(cleaned_input)
                df_kept, df_removed = filter_zero_prices(calculated_df)

                # === PEMBERSIHAN BARIS & KOLOM KOSONG ===
                # 1. Menentukan kolom wajib yang tidak boleh kosong (blank)
                kolom_wajib = ["Total Harga", "@Harga", "Barang & Jasa", "Nama Pelanggan"]

                # Memastikan kolom tersebut benar-benar ada di dalam dataframe yang sedang aktif
                kolom_tersedia = [col for col in kolom_wajib if col in df_kept.columns]

                # 2. Menghapus (drop) baris secara permanen jika ada yang kosong pada kolom wajib
                if kolom_tersedia:
                    df_kept = df_kept.dropna(subset=kolom_tersedia)

                empty_columns = [
                    column
                    for column in df_kept.columns
                    if column != REVIEW_STATUS_COLUMN and df_kept[column].isna().all()
                ]
                df_kept = df_kept.drop(columns=empty_columns)
                if REVIEW_STATUS_COLUMN not in df_kept.columns:
                    st.error(
                        "Format pemrosesan tidak sesuai: kolom 'Status' "
                        "tidak tersedia untuk validasi data."
                    )
                    st.stop()

                customer_summary = generate_rekap_pelanggan(df_kept)
                product_summary = generate_rekap_barang(df_kept)

                review_records = []
                validation_records = [
                    {
                        "Pemeriksaan": "Struktur kolom wajib",
                        "Status": "LULUS",
                        "Keterangan": "Semua kolom wajib tersedia.",
                    }
                ]
                review_mask = pd.Series(False, index=df_kept.index)
                review_reasons = {}

                volume_review = df_kept[REVIEW_STATUS_COLUMN].eq(REVIEW_STATUS)
                review_mask |= volume_review
                for index in df_kept.index[volume_review]:
                    review_reasons.setdefault(index, []).append(
                        "Kuantitas atau rasio tidak valid, atau aturan produk PCS "
                        "tidak tersedia."
                    )
                validation_records.append(
                    {
                        "Pemeriksaan": "Perhitungan Bata (M3)",
                        "Status": "PERLU DITINJAU" if volume_review.any() else "LULUS",
                        "Keterangan": (
                            f"{int(volume_review.sum())} baris perlu ditinjau."
                            if volume_review.any()
                            else "Semua baris berhasil dihitung."
                        ),
                    }
                )

                invalid_price_rows = pd.Series(False, index=df_kept.index)
                for price_column in ("Total Harga", "@Harga"):
                    numeric_prices = pd.to_numeric(
                        df_kept[price_column], errors="coerce"
                    )
                    invalid = numeric_prices.isna()
                    invalid_price_rows |= invalid
                    for index in df_kept.index[invalid]:
                        review_reasons.setdefault(index, []).append(
                            f"Nilai '{price_column}' kosong atau bukan angka."
                        )
                review_mask |= invalid_price_rows
                validation_records.append(
                    {
                        "Pemeriksaan": "Validitas harga",
                        "Status": (
                            "PERLU DITINJAU"
                            if invalid_price_rows.any()
                            else "LULUS"
                        ),
                        "Keterangan": (
                            f"{int(invalid_price_rows.sum())} baris memiliki "
                            "nilai harga kosong atau tidak valid."
                            if invalid_price_rows.any()
                            else "Semua nilai harga dapat dibaca sebagai angka."
                        ),
                    }
                )

                invalid_quantity = pd.to_numeric(
                    df_kept["Kuantitas"], errors="coerce"
                ).isna()
                validation_records.append(
                    {
                        "Pemeriksaan": "Validitas kuantitas",
                        "Status": (
                            "PERLU DITINJAU" if invalid_quantity.any() else "LULUS"
                        ),
                        "Keterangan": (
                            f"{int(invalid_quantity.sum())} baris memiliki "
                            "kuantitas kosong atau tidak valid."
                            if invalid_quantity.any()
                            else "Semua kuantitas dapat dibaca sebagai angka."
                        ),
                    }
                )

                for index in df_kept.index[review_mask]:
                    record = df_kept.loc[index].to_dict()
                    record["Nomor baris sumber"] = int(index) + 1
                    record["Alasan"] = " ".join(review_reasons[index])
                    review_records.append(record)
                review_df = pd.DataFrame(review_records)
                validation_df = pd.DataFrame(validation_records)

                removed_log = df_removed.copy()
                removed_log.insert(
                    0, "Nomor baris sumber", [int(index) + 1 for index in df_removed.index]
                )

                log_df = pd.DataFrame(
                    [
                        {"Informasi": "Nama file sumber", "Nilai": uploaded_file.name},
                        {
                            "Informasi": "Waktu proses",
                            "Nilai": datetime.now().isoformat(timespec="seconds"),
                        },
                        {"Informasi": "Perusahaan", "Nilai": selected_company},
                        {"Informasi": "Jumlah baris awal", "Nilai": initial_rows},
                        {
                            "Informasi": "Baris kosong dihapus",
                            "Nilai": blank_rows_removed,
                        },
                        {
                            "Informasi": "Baris harga nol dihapus",
                            "Nilai": len(df_removed),
                        },
                        {
                            "Informasi": "Jumlah baris akhir",
                            "Nilai": len(df_kept),
                        },
                        {
                            "Informasi": "Baris perlu ditinjau",
                            "Nilai": len(review_df),
                        },
                        {
                            "Informasi": "Detail baris harga nol",
                            "Nilai": (
                                removed_log.to_json(orient="records", force_ascii=False)
                                if not removed_log.empty
                                else "Tidak ada"
                            ),
                        },
                        {"Informasi": "Status proses", "Nilai": "Selesai"},
                    ]
                )
                info_df = pd.DataFrame(
                    [
                        {"Informasi": "Nama aplikasi", "Nilai": "Otomatisasi Surat Jalan"},
                        {"Informasi": "Perusahaan", "Nilai": selected_company},
                        {"Informasi": "Nama file sumber", "Nilai": uploaded_file.name},
                        {
                            "Informasi": "Waktu pengolahan",
                            "Nilai": datetime.now().isoformat(timespec="seconds"),
                        },
                        {
                            "Informasi": "Versi aturan bisnis",
                            "Nilai": "1.0",
                        },
                        {
                            "Informasi": "Rumus Bata (M3)",
                            "Nilai": "PCS: Kuantitas / rasio; selain PCS: Kuantitas.",
                        },
                        {
                            "Informasi": "Keterangan",
                            "Nilai": (
                                "Baris berstatus PERLU_DITINJAU memerlukan pemeriksaan "
                                "manual."
                            ),
                        },
                    ]
                )

                st.session_state["processed_result"] = {
                    "source_file": uploaded_file.name,
                    "company": selected_company,
                    "initial_rows": initial_rows,
                    "blank_rows_removed": blank_rows_removed,
                    "price_rows_removed": len(df_removed),
                    "cleaned_rows": len(df_kept),
                    "review_rows": len(review_df),
                    "df_kept": df_kept,
                    "df_removed": df_removed,
                    "customer_summary": customer_summary,
                    "product_summary": product_summary,
                    "review_df": review_df,
                    "validation_df": validation_df,
                    "log_df": log_df,
                    "info_df": info_df,
                }
            except (ValueError, TypeError, KeyError) as error:
                st.error(f"Data tidak dapat diproses: {error}")

result = st.session_state.get("processed_result")
if result is not None and (
    uploaded_file is None
    or result["source_file"] != uploaded_file.name
    or result["company"] != selected_company
):
    result = None

if result is not None:
    st.subheader("Hasil pengolahan")
    df_bersih = result["df_kept"].copy()
    kolom_sales = "Nama Sales"
    sales_terpilih = []

    if kolom_sales in df_bersih.columns:
        daftar_sales = sorted(
            df_bersih[kolom_sales].dropna().unique().tolist()
        )
        sales_terpilih = st.multiselect(
            "Filter berdasarkan Nama Sales (Kosongkan untuk mengekspor semua sales):",
            options=daftar_sales,
            help=(
                "Data Bersih dan seluruh rekap di bawahnya akan menyesuaikan "
                "dengan sales yang dipilih."
            ),
        )
        if sales_terpilih:
            df_bersih = df_bersih[
                df_bersih[kolom_sales].isin(sales_terpilih)
            ]

    sales_scope = repr(tuple(sales_terpilih))
    customer_column = "Nama Pelanggan"
    product_column = "Barang & Jasa"
    pelanggan_tersedia = (
        sorted(df_bersih[customer_column].dropna().unique().tolist())
        if customer_column in df_bersih.columns
        else []
    )
    metric_columns = st.columns(4)
    metric_columns[0].metric("Baris awal", f"{result['initial_rows']:,}")
    metric_columns[1].metric(
        "Baris dihapus",
        f"{result['blank_rows_removed'] + result['price_rows_removed']:,}",
    )
    metric_columns[2].metric("Baris data bersih", f"{len(df_bersih):,}")
    metric_columns[3].metric("Perlu ditinjau", f"{result['review_rows']:,}")

    # Pengecekan final sebelum render ke layar
    ada_axel_final = (
        df_bersih["Barang & Jasa"]
        .astype(str)
        .str.contains("AXEL", case=False, na=False)
        .any()
    )
    df_tampilan = df_bersih.copy()
    if not ada_axel_final and "Bata (M3)" in df_tampilan.columns:
        df_tampilan = df_tampilan.drop(columns=["Bata (M3)"])

    total_quantity = pd.to_numeric(
        df_tampilan["Kuantitas"], errors="coerce"
    ).sum()
    teks_kuantitas = f"Total kuantitas: **{total_quantity:,.2f}**"
    if "Bata (M3)" in df_tampilan.columns:
        total_volume = pd.to_numeric(
            df_tampilan["Bata (M3)"], errors="coerce"
        ).sum()
        st.markdown(
            f"{teks_kuantitas} &middot; Total Bata (M3) yang dapat dihitung: "
            f"**{total_volume:,.2f}**"
        )
    else:
        st.markdown(teks_kuantitas)
    if result["review_rows"]:
        st.warning("Validasi hasil: ada data yang perlu ditinjau.")
    else:
        st.success("Validasi hasil: seluruh data lolos pemeriksaan.")

    if product_column in df_bersih.columns:
        df_tampilan = df_tampilan.copy()
        if "Bata (M3)" in df_tampilan.columns:
            mask = df_tampilan[product_column].str.contains(
                "PALLET", case=False, na=False
            )
            df_tampilan.loc[mask, "Bata (M3)"] = df_tampilan.loc[
                mask, "Kuantitas"
            ]
        st.dataframe(df_tampilan, width="stretch")
    else:
        st.dataframe(df_tampilan, width="stretch")

    st.write("---")
    st.subheader("Preview Pelanggan")

    jumlah_grup = st.number_input(
        "Jumlah Grup Klasifikasi", min_value=1, max_value=10, value=1, step=1
    )
    kumpulan_tabel_rekap = []

    for i in range(jumlah_grup):
        st.write(f"**Pengaturan Grup {i + 1}**")
        col1, col2 = st.columns(2)

        daftar_barang_asli = (
            sorted(df_bersih[product_column].dropna().unique().tolist())
            if product_column in df_bersih.columns
            else []
        )
        kata_kunci = sorted(
            list(
                set(
                    [
                        str(barang).split()[0].upper()
                        for barang in daftar_barang_asli
                        if str(barang).strip()
                    ]
                )
            )
        )
        daftar_shortcut = [f"SEMUA {kata}" for kata in kata_kunci]
        opsi_barang_lengkap = daftar_shortcut + daftar_barang_asli

        with col1:
            pelanggan_terpilih = st.multiselect(
                "Pilih Nama Pelanggan",
                options=pelanggan_tersedia,
                key=f"pelanggan_{sales_scope}_{i}",
            )

        with col2:
            barang_terpilih = st.multiselect(
                "Pilih Barang",
                options=opsi_barang_lengkap,
                key=f"barang_{sales_scope}_{i}",
            )

        # SISIPAN PERBAIKAN UI: Pindahkan Paste dan Foto ke dalam Expander
        with st.expander("Input Klasifikasi Pelanggan Lanjutan"):
            col_paste, col_foto = st.columns(2)

            with col_paste:
                input_massal = st.text_area(
                    "Paste Daftar Nama Pelanggan (Pisahkan dengan Enter)",
                    key=f"paste_pelanggan_{i}",
                    height=100,
                )

            with col_foto:
                upload_foto = st.file_uploader(
                    "Unggah Foto Tabel Pelanggan (JPG/PNG)",
                    type=["png", "jpg", "jpeg"],
                    key=f"foto_pelanggan_{i}",
                )

                teks_dari_foto = ""
                if upload_foto is not None:
                    try:
                        gambar = Image.open(upload_foto)
                        teks_dari_foto = pytesseract.image_to_string(gambar)
                        st.caption("Foto berhasil dipindai.")
                    except (
                        OSError,
                        UnidentifiedImageError,
                        pytesseract.TesseractNotFoundError,
                        pytesseract.TesseractError,
                    ) as error:
                        st.error(
                            "Gagal membaca foto. Pastikan file gambar valid dan "
                            f"Tesseract OCR terinstal di sistem: {error}"
                        )

        filter_pelanggan_final = list(pelanggan_terpilih)
        # SISIPAN BARU: Menggabungkan hasil paste dengan hasil multiselect
        if input_massal:
            daftar_paste = [
                nama.strip().upper()
                for nama in input_massal.split("\n")
                if nama.strip()
            ]
            pelanggan_tersedia_asli = (
                df_bersih["Nama Pelanggan"].dropna().unique().tolist()
            )

            for nama_asli in pelanggan_tersedia_asli:
                if str(nama_asli).strip().upper() in daftar_paste:
                    filter_pelanggan_final.append(nama_asli)

            filter_pelanggan_final = list(set(filter_pelanggan_final))

        # SISIPAN BARU: Memproses hasil teks dari foto OCR
        if teks_dari_foto:
            daftar_foto = [
                nama.strip().upper()
                for nama in teks_dari_foto.split("\n")
                if len(nama.strip()) > 2
            ]
            pelanggan_tersedia_asli = (
                df_bersih["Nama Pelanggan"].dropna().unique().tolist()
            )

            for nama_asli in pelanggan_tersedia_asli:
                if any(
                    str(nama_asli).strip().upper() in baris_ocr
                    for baris_ocr in daftar_foto
                ):
                    filter_pelanggan_final.append(nama_asli)

            filter_pelanggan_final = list(set(filter_pelanggan_final))

        df_grup = df_bersih.copy()
        if filter_pelanggan_final:
            df_grup = df_grup[
                df_grup["Nama Pelanggan"].isin(filter_pelanggan_final)
            ]
        if barang_terpilih:
            filter_barang_final = []
            for pilihan in barang_terpilih:
                if str(pilihan).startswith("SEMUA "):
                    kata = pilihan.replace("SEMUA ", "").strip()
                    filter_barang_final.extend(
                        [
                            barang
                            for barang in daftar_barang_asli
                            if str(barang).upper().startswith(kata)
                        ]
                    )
                else:
                    filter_barang_final.append(pilihan)

            filter_barang_final = list(set(filter_barang_final))
            df_grup = df_grup[
                df_grup["Barang & Jasa"].isin(filter_barang_final)
            ]

        if not df_grup.empty:
            baris_tampilan = []
            kuantitas = pd.to_numeric(df_grup["Kuantitas"], errors="coerce")
            grand_total = kuantitas.sum()
            data_rekap = df_grup.assign(Kuantitas=kuantitas)
            subtotal_pelanggan = data_rekap.groupby(
                "Nama Pelanggan", dropna=False, sort=False
            )["Kuantitas"].sum()

            for pelanggan, data_pelanggan in data_rekap.groupby(
                "Nama Pelanggan", dropna=False, sort=False
            ):
                subtotal = (
                    subtotal_pelanggan.loc[pelanggan]
                    if not pd.isna(pelanggan)
                    else data_pelanggan["Kuantitas"].sum()
                )
                baris_tampilan.append(
                    {
                        "Nama Pelanggan": pelanggan,
                        "Barang & Jasa": "TOTAL",
                        "Kuantitas": subtotal,
                    }
                )
                kuantitas_produk = data_pelanggan.groupby(
                    "Barang & Jasa", dropna=False, sort=False
                )["Kuantitas"].sum()
                for produk, kuantitas_produk_pelanggan in kuantitas_produk.items():
                    baris_tampilan.append(
                        {
                            "Nama Pelanggan": "",
                            "Barang & Jasa": f"  {produk}",
                            "Kuantitas": kuantitas_produk_pelanggan,
                        }
                    )

            baris_tampilan.append(
                {
                    "Nama Pelanggan": "Grand Total",
                    "Barang & Jasa": "",
                    "Kuantitas": grand_total,
                }
            )
            df_preview = pd.DataFrame(baris_tampilan)
            kumpulan_tabel_rekap.append(pd.DataFrame(df_preview))
            st.dataframe(df_preview, hide_index=True)

    # --- LOGIKA PENYESUAIAN AXEL FINAL ---
    ada_axel_final = (
        df_bersih["Barang & Jasa"]
        .astype(str)
        .str.contains("AXEL", case=False, na=False)
        .any()
    )
    if not ada_axel_final:
        if "Bata (M3)" in df_bersih.columns:
            df_bersih = df_bersih.drop(columns=["Bata (M3)"])
        if "Bata (M3)" in result["product_summary"].columns:
            result["product_summary"] = result["product_summary"].drop(
                columns=["Bata (M3)"]
            )

    # Buang kolom dummy Alamat khusus Sangnila (jika ada)
    if (
        selected_company == "CV Sangnila Tirta Utama"
        and "Alamat pengiriman" in df_bersih.columns
    ):
        df_bersih = df_bersih.drop(columns=["Alamat pengiriman"])

    dict_of_dfs = {
        "Rekap_Pelanggan": kumpulan_tabel_rekap,
        "Data_Bersih": df_bersih,
        "Rekap_Barang": result["product_summary"],
        "Validasi_Data": result["validation_df"],
        "Perlu_Ditinjau": result["review_df"],
        "Log_Pengolahan": result["log_df"],
        "Info": result["info_df"],
    }
    try:
        # Hasilkan file Excel dalam bentuk BytesIO
        excel_data = export_to_excel(dict_of_dfs)

        # Tombol unduh bawaan Streamlit
        st.download_button(
            label="Unduh hasil Excel",
            data=excel_data,
            file_name="hasil_pengolahan_surat_jalan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")

st.caption(
    "File Excel yang didukung: .xlsx dan .xlsm. File sumber tidak diubah."
)
