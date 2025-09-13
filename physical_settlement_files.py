import math
import os
import pandas as pd
import cons_header

def safe_float(val):
    """Convert to float safely, treating blanks/NA as 0.0"""
    if pd.isna(val):
        return 0
    try:
        num = float(val)
        # Round .5 and above up, below .5 down
        return math.floor(num + 0.5)
    except Exception:
        return 0
    
def read_file(filepath: str, header: int = 0, custom_header: list = None) -> pd.DataFrame:
    """
    Reads a file (CSV, XLS, XLSX) into a pandas DataFrame.
    
    Args:
        filepath (str): Path to the input file
        header (int): Row number to use as column names (default=0)
        custom_header (list): If provided, select only these columns after loading
    
    Returns:
        pd.DataFrame: Loaded data
    """
    # Load with original headers from file
    if filepath.lower().endswith(".csv"):
        df = pd.read_csv(filepath, header=header)
    elif filepath.lower().endswith(".xlsx"):
        df = pd.read_excel(filepath, engine="openpyxl", header=header)
    elif filepath.lower().endswith(".xls"):
        df = pd.read_excel(filepath, engine="xlrd", header=header)
    else:
        raise ValueError("Unsupported file format. Supported: .csv, .xls, .xlsx")
    
    # Strip spaces from column headers
    df.columns = df.columns.str.strip()

    # If custom_header is provided, select only those columns
    if custom_header:
        custom_header = [col.strip() for col in custom_header]
        missing = [col for col in custom_header if col not in df.columns]
        if missing:
            raise ValueError(f"These columns are missing in file: {missing}")
        df = df[custom_header]
    
    return df

# MOMF_PATH = r"D:\Ranjan sir\physical settlement files\MOMF_Obligation_Physical Settlement_28082025.xlsx"
# OBLIGATION_PATH = r"D:\Ranjan sir\physical settlement files\Obligation_NCL_FO_FOPHY_CM_90123_20250828_F_0000.csv"
# STAMP_DUTY_PATH = r"D:\Ranjan sir\physical settlement files\StampDuty_NCL_FO_FOPHY_CM_90123_20250828_F_0000.csv"
# STT_PATH = r"D:\Ranjan sir\physical settlement files\STT_NCL_FO_FOPHY_CM_90123_20250828_F_0000.csv"

# momf_df = read_file(MOMF_PATH)
# obligation_df = read_file(OBLIGATION_PATH)
# stamp_duty_df = read_file(STAMP_DUTY_PATH)
# stt_df = read_file(STT_PATH)

def build_dict(file_path: str, key_cols: list, value_cols: dict, filter_col: str = None, filter_val=None) -> dict:
    """
    Reads CSV, applies optional filter, builds dictionary for manual updates.

    Args:
        file_path (str): Path to CSV file
        key_cols (list): Column names to form dictionary key (tuple if >1)
        value_cols (dict): { "Output Column Name": "Source Column Name" }
        filter_col (str, optional): Column name for filter
        filter_val (any, optional): Value to filter rows

    Returns:
        dict: { (keys): { "Output Column": value, ... } }
    """
    df = pd.read_csv(file_path, header=0)
    df.columns = df.columns.str.strip()

    # Apply filter if provided
    if filter_col and filter_val is not None:
        df = df[df[filter_col] == filter_val]

    lookup_dict = {}
    for _, row in df.iterrows():
        key = tuple(str(row[col]).strip() for col in key_cols)

        value_dict = {}
        for out_col, src_col in value_cols.items():
            val = row[src_col]

            # Replace NaN/None with 0
            if pd.isna(val):
                val = ''

            value_dict[out_col] = safe_float(val)

        lookup_dict[key] = value_dict

    return lookup_dict


def segregate_excel_by_column(excel_path: str,
                                output_path: str,
                                column_name: str = "BrkrOrCtdnPtcptId", 
                                custom_header: list = None,
                                update_dicts: list = None):
    """
    Segregates Excel data by unique values of a given column 
    and writes into multiple sheets.
    
    Args:
        excel_path (str): Path to Excel file
        output_path (str): Path to save segregated Excel file
        column_name (str): Column to group by (default: BrkrOrCtdnPtcptId)
    """
    df = read_file(excel_path, custom_header=custom_header)  # use your read function
    # print("name:", column_name)
    # print("Unique BrkrOrCtdnPtcptId:", df[column_name].unique())

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in file. Found: {df.columns.tolist()}")

    # Add extra columns at the end in the specified order
    for col in cons_header.EXTRA_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA #  pd.NA  preferred for empty cells in pandas
    
    # Replace blanks/NaN in grouping column with "Blank"
    df[column_name] = df[column_name].fillna("").astype(str).str.strip()
    df[column_name] = df[column_name].replace({"": "Blank"})

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for key, subset in df.groupby(column_name):
            subset = subset.copy()
            # Update values from each update_dict if key matches
            for i, row in subset.iterrows():
                val = row["FinInstrmId"]
                if pd.isna(val):        # NaN in pandas
                    val_float = None    # or leave as None
                elif str(val).strip() == "":  # Blank string
                    val_float = None
                else:
                    val_float = float(val)
                    val_str = str(val_float) # convert float to string

                match_key = (row["BrkrOrCtdnPtcptId"], row["TckrSymb"], val_str)

                # print("Matching Key:", match_key)
                for update_dict in update_dicts:
                    if match_key in update_dict:
                        for col, val in update_dict[match_key].items():
                            subset.at[i, col] = val  # just update directly

            # now i want to access here df and access some columns and do some calculations and update those columns
            # Do calculations for each row in the sheet
            for i, row in subset.iterrows():
                cmltv_buy   = safe_float(row.get("CmltvBuyAmt"))
                buy_stt     = safe_float(row.get("Buy STT"))
                buy_stamp   = safe_float(row.get("Buy Stamp Duty"))

                cmltv_sell  = safe_float(row.get("CmltvSellAmt"))
                sell_stt    = safe_float(row.get("Sell STT"))

                buy_payable = cmltv_buy + buy_stt + buy_stamp
                sell_receivable = cmltv_sell - sell_stt
                net_receivable_payable = sell_receivable - buy_payable

                subset.at[i, "Buy Payable Amount"] = buy_payable
                subset.at[i, "Sell Receivable Amount"] = sell_receivable
                subset.at[i, "Net Receivable \\ Payable"] = net_receivable_payable
            
            # ======================
            # Add Total row for "Net Receivable \ Payable"
            # ======================
            if "Net Receivable \\ Payable" in subset.columns:
                total_value = subset["Net Receivable \\ Payable"].apply(safe_float).sum()

                total_row = {col: "" for col in subset.columns}
                cols = list(subset.columns)
                net_idx = cols.index("Net Receivable \\ Payable")

                if net_idx > 0:
                    total_row[cols[net_idx - 1]] = "Total"   # left side of sum
                total_row["Net Receivable \\ Payable"] = total_value

                subset = pd.concat([subset, pd.DataFrame([total_row])], ignore_index=True)

            safe_sheet_name = str(key)[:31] or "Blank"  # Excel max 31 chars
            subset.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    
    # print(f"✅ Segregated file saved: {output_path}")

# Buy Payable Amount =Z4+AC4+AF4  CmltvBuyAmt+Buy STT+Buy Stamp Duty
# Sell Receivable Amount = AA13-AD13 CmltvSellAmt-Sell STT
# Net Receivable \ Payable =AH4-AG4 Sell Receivable Amount-Buy Payable Amount


# Build STT dictionary
# stt_dict = build_dict(
#     file_path=STT_PATH,
#     key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
#     value_cols={
#         "Buy STT": "BuyDelvryTtlTaxs",
#         "Sell STT": "SellDelvryTtlTaxs"
#     },
#     filter_col="RptHdr",
#     filter_val=40
# )

# stamp_duty_dict = build_dict(file_path=STAMP_DUTY_PATH,
#     key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
#     value_cols={
#         # "Sell Stamp Duty": "",
#         "Buy Stamp Duty": "BuyDlvryStmpDty"
#     },
#     filter_col="RptHdr",
#     filter_val=40
# )

# # Segregate obligation and update Buy/Sell STT from dictionary
# segregate_excel_by_column(
#     excel_path=OBLIGATION_PATH,
#     output_path=r"D:\Ranjan sir\physical settlement files\Output\Segregated_Output.xlsx",
#     column_name="BrkrOrCtdnPtcptId",
#     custom_header=cons_header.OBLIGATION_HEADER,
#     update_dicts=[stt_dict, stamp_duty_dict]
# )

