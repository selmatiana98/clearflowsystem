from pathlib import Path
from typing import BinaryIO

import pandas as pd

from config.perusahaan import KONFIGURASI_PERUSAHAAN


HEADER_SCAN_ROWS = 15


def read_excel_file(
    file_path: str | Path | BinaryIO,
    required_columns: list[str] | None = None,
) -> tuple[pd.DataFrame, str]:
    """Read the first sheet with a recognizable header and normalize its columns."""
    use_default_columns = required_columns is None
    if required_columns is None:
        required_columns = [
            "Nama Pelanggan",
            "Barang & Jasa",
            "Satuan",
            "Kuantitas",
            "Rasio Satuan #2 Barang & Jasa",
            "Alamat pengiriman",
            "Total Harga",
            "@Harga",
        ]
    column_mapping = {
        "Pelanggan": "Nama Pelanggan",
        "Nama Barang": "Barang & Jasa",
        "Alamat Pengiriman": "Alamat pengiriman",
    }

    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        if not sheet_names:
            raise ValueError("Workbook tidak memiliki sheet.")

        for sheet_name in sheet_names:
            preview = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                nrows=HEADER_SCAN_ROWS,
            )
            header_row = None
            for row_index, row in preview.iterrows():
                values = [
                    str(value).strip().casefold()
                    for value in row
                    if pd.notna(value)
                ]
                if use_default_columns and (
                    any("kuantitas" in value for value in values)
                    and any("satuan" in value for value in values)
                ):
                    header_row = row_index
                    break
                if not use_default_columns and all(
                    column.strip().casefold() in values
                    for column in required_columns
                ):
                    header_row = row_index
                    break

            if header_row is None:
                continue

            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
            df.columns = df.columns.astype(str).str.strip()
            if use_default_columns:
                df = df.rename(columns=column_mapping)
                normalized_columns = {
                    **{
                        source.casefold(): target
                        for source, target in column_mapping.items()
                    },
                    **{column.casefold(): column for column in required_columns},
                }
            else:
                normalized_columns = {
                    column.casefold(): column for column in required_columns
                }
            df.columns = [
                normalized_columns.get(column.casefold(), column)
                for column in df.columns
            ]
            if all(
                list(df.columns).count(column) == 1 for column in required_columns
            ):
                return df, sheet_name

        sheets = ", ".join(sheet_names)
        raise ValueError(
            f"Header dengan seluruh kolom wajib tidak ditemukan dalam {HEADER_SCAN_ROWS} "
            f"baris pertama sheet mana pun ({sheets}). "
            f"Kolom wajib: {', '.join(required_columns)}."
        )
