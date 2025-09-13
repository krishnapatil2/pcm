import gzip
import csv
import pandas as pd

def read_sec_pledge_gz(file_path, marker='GSEC', encoding='utf-8', sample_size=8192):
    """
    Robustly read a .csv.gz SEC_PLEDGE file:
    - sniff delimiter
    - parse with csv.reader (handles quoted fields & embedded newlines)
    - find the row where first column == marker (case-insensitive)
    - treat the next CSV row as header and return a DataFrame + lookup dict
    """
    # 1) sniff delimiter from a small sample
    with gzip.open(file_path, "rt", encoding=encoding, errors="ignore", newline='') as f:
        sample = f.read(sample_size)
    try:
        dialect = csv.Sniffer().sniff(sample)
        delim = dialect.delimiter
    except Exception:
        delim = ','

    # 2) parse entire file as CSV (logical rows)
    with gzip.open(file_path, "rt", encoding=encoding, errors="ignore", newline='') as f:
        reader = csv.reader(f, delimiter=delim)
        rows = list(reader)

    # Diagnostics (optional - helpful while debugging)
    # print(f"Physical/text sample size: {len(sample)} chars, parsed CSV rows: {len(rows)}, guessed delimiter: '{delim}'")

    # 3) find marker row (first column equals marker)
    marker_row = None
    for i, row in enumerate(rows):
        if len(row) == 0:
            continue
        first_col = row[0].strip().upper()
        if first_col == marker.upper():
            marker_row = i
            break

    if marker_row is None:
        raise ValueError(f"Marker '{marker}' not found in first column of any CSV row.")

    header_idx = marker_row + 1
    if header_idx >= len(rows):
        raise ValueError(f"Marker found at CSV row {marker_row} but there is no following header row (rows={len(rows)}).")

    header = [h.strip() for h in rows[header_idx]]
    data_rows = rows[header_idx + 1 :]

    # 4) Build DataFrame from parsed rows
    df = pd.DataFrame(data_rows, columns=header)

    # 5) Clean column names
    df.columns = df.columns.str.strip()

    # 6) Optional: try to coerce known numeric columns to numeric (safe)
    for col in ["GROSS VALUE", "HAIRCUT"]:
        if col in df.columns:
            # remove thousand-separators and convert; non-numeric -> NaN
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "", regex=False)
                .replace("", pd.NA)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 7) Build lookup dict
    sec_pledge_lookup = {}
    for _, row in df.iterrows():
        client_code = str(row.get("Client/CP code", "")).strip()
        isin = str(row.get("ISIN", "")).strip()
        gross_value = row.get("GROSS VALUE", pd.NA)
        haircut = row.get("HAIRCUT", pd.NA)

        if not client_code or not isin or client_code.upper() in ("", "NONE", "N/A"):
            continue

        key = f"{client_code}-{isin}"
        sec_pledge_lookup[key] = {"GROSS VALUE": gross_value, "HAIRCUT": haircut}

    return df, sec_pledge_lookup


file = r"C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\F_90123_SEC_PLEDGE_11092025_02.csv.gz"
df8, _sec_pledge_lookup = read_sec_pledge_gz(file)
print("Rows (CSV logical):", len(df8))
print("Columns:", df8.columns.tolist())
print("Sample lookup entries:", list(_sec_pledge_lookup.items())[:5])
