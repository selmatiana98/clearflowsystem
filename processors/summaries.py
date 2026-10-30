import pandas as pd

def generate_rekap_pelanggan(df):
    try:
        kolom_agregasi = {}
        for col in ["Kuantitas", "Total Harga", "@Harga", "Bata (M3)"]:
            if col in df.columns:
                kolom_agregasi[col] = "sum"
                
        rekap = df.groupby("Nama Pelanggan", dropna=False).agg(kolom_agregasi).reset_index()
        return rekap
    except Exception as e:
        raise ValueError(f"kolom bermasalah pada rekap pelanggan: {e}")

def generate_rekap_barang(df):
    try:
        kolom_agregasi = {}
        for col in ["Kuantitas", "Total Harga", "@Harga", "Bata (M3)"]:
            if col in df.columns:
                kolom_agregasi[col] = "sum"
                
        rekap = df.groupby(["Barang & Jasa", "Satuan"], dropna=False).agg(kolom_agregasi).reset_index()
        return rekap
    except Exception as e:
        raise ValueError(f"kolom bermasalah pada rekap barang: {e}")