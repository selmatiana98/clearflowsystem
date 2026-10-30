import math
import warnings

import pandas as pd


AXEL_AAC_FALLBACK_RATIOS = {
    "AXEL AAC 10X20X40CM": 83,
    "AXEL AAC 10X20X50CM": 100,
    "AXEL AAC 10X20X58CM": 83,
    "AXEL AAC 10X20X60CM": 83,
    "AXEL AAC 7,5X20X40CM": 111,
    "AXEL AAC 7,5X20X50CM": 111,
    "AXEL AAC 7,5X20X60CM": 111,
}
REVIEW_STATUS = "PERLU_DITINJAU"
REVIEW_STATUS_COLUMN = "Status"


def calculate_bata_m3(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate Bata (M3), using ratios for AXEL AAC PCS and quantity otherwise."""
    # --- SENSOR AXEL ---
    ada_axel = df['Barang & Jasa'].astype(str).str.contains('AXEL', case=False, na=False).any()
    if not ada_axel:
        df_bypass = df.copy()
        # Cegah error di app.py dengan memberikan status default
        if REVIEW_STATUS_COLUMN not in df_bypass.columns:
            df_bypass[REVIEW_STATUS_COLUMN] = "LULUS"
        return df_bypass
    # -------------------

    required_columns = (
        "Satuan",
        "Kuantitas",
        "Barang & Jasa",
        "Rasio Satuan #2 Barang & Jasa",
        "Alamat pengiriman",
    )
    missing_columns = [column for column in required_columns if column not in df.columns]
    duplicate_columns = [
        column
        for column in required_columns + ("Bata (M3)",)
        if list(df.columns).count(column) > 1
    ]
    conflicting_columns = [REVIEW_STATUS_COLUMN] if REVIEW_STATUS_COLUMN in df else []
    if missing_columns or duplicate_columns or conflicting_columns:
        problems = []
        if missing_columns:
            problems.append(f"kolom tidak ditemukan: {', '.join(missing_columns)}")
        if duplicate_columns:
            problems.append(f"nama kolom ambigu: {', '.join(duplicate_columns)}")
        if conflicting_columns:
            problems.append(
                f"kolom keluaran sudah ada: {', '.join(conflicting_columns)}"
            )
        message = (
            "Tidak dapat menghitung Bata (M3) atau menempatkan kolom pada posisi "
            "yang ditentukan; " + "; ".join(problems)
        )
        warnings.warn(message, RuntimeWarning, stacklevel=2)
        raise ValueError(message)

    result = df.copy()
    quantities = pd.to_numeric(result["Kuantitas"], errors="coerce")
    accurate_ratios = pd.to_numeric(
        result["Rasio Satuan #2 Barang & Jasa"], errors="coerce"
    )
    ada_axel = (
        result["Barang & Jasa"]
        .astype(str)
        .str.contains("AXEL", case=False, na=False)
        .any()
    )
    calculated_values = []
    row_statuses = []

    for unit, quantity, product, accurate_ratio in zip(
        result["Satuan"],
        quantities,
        result["Barang & Jasa"],
        accurate_ratios,
    ):
        if pd.isna(quantity) or not math.isfinite(quantity):
            calculated_values.append(None)
            row_statuses.append(REVIEW_STATUS)
            continue

        if not ada_axel:
            calculated_values.append(quantity)
            row_statuses.append(None)
            continue

        if (
            not isinstance(unit, str)
            or unit.strip().casefold() != "pcs"
            or not isinstance(product, str)
            or not product.strip().upper().startswith("AXEL AAC")
        ):
            calculated_values.append(quantity)
            row_statuses.append(None)
            continue

        ratio = accurate_ratio
        if pd.isna(ratio) or not math.isfinite(ratio) or ratio <= 0:
            normalized_product = (
                product.strip().upper() if isinstance(product, str) else ""
            )
            ratio = AXEL_AAC_FALLBACK_RATIOS.get(normalized_product)

        if ratio is None:
            calculated_values.append(None)
            row_statuses.append(REVIEW_STATUS)
        else:
            volume = quantity / ratio
            if math.isfinite(volume):
                calculated_values.append(volume)
                row_statuses.append(None)
            else:
                calculated_values.append(None)
                row_statuses.append(REVIEW_STATUS)

    result = result.drop(columns=["Bata (M3)"], errors="ignore")
    insert_at = result.columns.get_loc("Rasio Satuan #2 Barang & Jasa") + 1
    result.insert(
        insert_at,
        "Bata (M3)",
        pd.Series(calculated_values, index=result.index, dtype=object),
    )
    result[REVIEW_STATUS_COLUMN] = pd.Series(
        row_statuses,
        index=result.index,
        dtype=object,
    )
    result["Bata (M3)"] = pd.to_numeric(result["Bata (M3)"], errors="coerce")
    result["Kuantitas"] = pd.to_numeric(result["Kuantitas"], errors="coerce")

    def enforce_bata_rule(row: pd.Series) -> object:
        barang = str(row.get("Barang & Jasa", "")).upper()
        kuantitas = row.get("Kuantitas", 0)
        bata_saat_ini = row.get("Bata (M3)")

        if "AXEL" not in barang:
            return kuantitas

        if (
            pd.isna(bata_saat_ini)
            or bata_saat_ini is None
            or str(bata_saat_ini).strip().lower() in ("", "none", "nan")
        ):
            return kuantitas

        return bata_saat_ini

    result["Bata (M3)"] = result.apply(enforce_bata_rule, axis=1)
    pallet_rows = result["Barang & Jasa"].astype("string").str.contains(
        "PALLET",
        case=False,
        na=False,
    )
    result.loc[pallet_rows, "Bata (M3)"] = result.loc[pallet_rows, "Kuantitas"]
    return result
