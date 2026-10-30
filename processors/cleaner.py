import pandas as pd


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Remove empty/auxiliary columns and format export-facing values."""
    mask = df.astype(str).apply(
        lambda column: column.str.contains(
            "ACCURATE|Tercetak|Halaman", case=False, na=False
        )
    ).any(axis=1)
    cleaned = (
        df.loc[~mask].dropna(axis=0, how="all").dropna(axis=1, how="all").copy()
    )
    columns_to_drop = [
        column
        for column in cleaned.columns
        if "status" in str(column).casefold()
        or "unnamed" in str(column).casefold()
    ]
    cleaned = cleaned.drop(columns=columns_to_drop)
    if "Tanggal" in cleaned.columns:
        formatted = pd.to_datetime(cleaned["Tanggal"], errors="coerce").dt.strftime(
            "%d %b %Y"
        )
        cleaned["Tanggal"] = "'" + formatted.fillna("")
    if "Bata (M3)" in cleaned.columns:
        cleaned["Bata (M3)"] = pd.to_numeric(
            cleaned["Bata (M3)"], errors="coerce"
        ).round(2)
    return cleaned
