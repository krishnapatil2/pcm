import math
import os
import pandas as pd
import cons_header
from openpyxl.styles import Border, Side

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
        # key = tuple(str(row[col]).strip() for col in key_cols)
        key = tuple(
                        str(int(row[col])) if col == "BrkrOrCtdnPtcptId" and isinstance(row[col], float) and row[col].is_integer()
                        else str(row[col]).strip()
                        for col in key_cols
                    )        

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

    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir:  # Only create directory if path has a directory component
        os.makedirs(output_dir, exist_ok=True)
    
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
                buystamp = row.get("Buy Stamp Duty")
                sellstamp = row.get("Sell Stamp Duty")

                # Handle missing stamp duties
                if pd.isna(buystamp) or str(buystamp).strip() == "":
                    subset.at[i, "Buy Stamp Duty"] = 0

                if pd.isna(sellstamp) or str(sellstamp).strip() == "":
                    subset.at[i, "Sell Stamp Duty"] = 0
                
                buy_stamp   = safe_float(subset.at[i, "Buy Stamp Duty"])
                sell_stamp  = safe_float(subset.at[i, "Sell Stamp Duty"])

                cmltv_buy   = safe_float(row.get("CmltvBuyAmt"))
                buy_stt     = safe_float(row.get("Buy STT"))

                cmltv_sell  = safe_float(row.get("CmltvSellAmt"))
                sell_stt    = safe_float(row.get("Sell STT"))

                buy_payable = cmltv_buy + buy_stt + buy_stamp
                sell_receivable = cmltv_sell - sell_stt - sell_stamp
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
            
            # Add borders to all cells in the sheet
            worksheet = writer.sheets[safe_sheet_name]
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Apply borders to all cells with data
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.border = thin_border
            
            # Auto-adjust column widths based on content
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # Set minimum width for EXTRA_COLUMNS and adjust based on content
                adjusted_width = max(max_length + 2, 12)  # Add padding and minimum width
                
                # Special handling for EXTRA_COLUMNS to ensure adequate width
                if column_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']:
                    # Check if this column contains EXTRA_COLUMNS data
                    header_cell = worksheet.cell(row=1, column=column[0].column)
                    if header_cell.value in cons_header.EXTRA_COLUMNS:
                        adjusted_width = max(adjusted_width, 18)  # Ensure extra columns have good width
                
                worksheet.column_dimensions[column_letter].width = adjusted_width