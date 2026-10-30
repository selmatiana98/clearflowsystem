import pandas as pd


def filter_zero_prices(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Separate rows with a numeric zero in either price column."""
    price_columns = ("Total Harga", "@Harga")
    missing_columns = [column for column in price_columns if column not in df.columns]
    duplicate_columns = [
        column
        for column in price_columns
        if list(df.columns).count(column) > 1
    ]
    if missing_columns or duplicate_columns:
        problems = []
        if missing_columns:
            problems.append(f"kolom tidak ditemukan: {', '.join(missing_columns)}")
        if duplicate_columns:
            problems.append(f"nama kolom ambigu: {', '.join(duplicate_columns)}")
        raise ValueError("Tidak dapat memfilter harga; " + "; ".join(problems))

    zero_price_masks = {
        column: pd.to_numeric(df[column], errors="coerce")
        .eq(0)
        .fillna(False)
        .tolist()
        for column in price_columns
    }
    removed_positions = [
        position
        for position in range(len(df))
        if any(zero_price_masks[column][position] for column in price_columns)
    ]
    removed_position_set = set(removed_positions)
    kept_positions = [
        position for position in range(len(df)) if position not in removed_position_set
    ]
    df_removed = df.iloc[removed_positions].copy()
    df_kept = df.iloc[kept_positions].copy()
    df_removed["Alasan penghapusan"] = [
        "; ".join(
            f"{column} = 0"
            for column in price_columns
            if zero_price_masks[column][position]
        )
        for position in removed_positions
    ]
    return df_kept, df_removed
