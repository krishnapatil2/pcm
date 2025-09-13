from CONSTANT_SEGREGATION import *
import pandas as pd
import os
import io
import warnings
warnings.filterwarnings("ignore")

def calculate_final_effective_value(gross_value, haircut):
    """
    Calculate Final Effective Value from Gross Value and Haircut.

    Parameters:
        gross_value (float or int): The gross value amount.
        haircut (float or int): The haircut percentage (can be 25 for 25% or 0.25 for 25%).

    Returns:
        float: Final Effective Value.
    """
    if gross_value is None:
        gross_value = 0
    if haircut is None:
        haircut = 0

    try:
        gross_value = float(gross_value)
        haircut = float(haircut)

        # Normalize haircut: if > 1, assume it's percentage (e.g., 25 -> 0.25)
        if haircut > 1:
            haircut = haircut / 100.0

        deduction = gross_value * haircut
        final_value = gross_value - deduction
        return round(final_value, 2)
    except Exception:
        return 0.0

def build_cp_lookup(sec_pledge_lookup):
    """
    Convert {cp-isin: {...}} → {cp: total_final_effective_value}
    """
    cp_lookup = {}

    for key, values in sec_pledge_lookup.items():
        # Split "CPCode-ISIN" → CPCode
        cp_code = key.split("-")[0].strip()

        gross_value = float(values.get("GROSS VALUE", 0))
        haircut = float(values.get("HAIRCUT", 0))

        final_effective_value = calculate_final_effective_value(gross_value, haircut)

        # Aggregate by CP code
        if cp_code in cp_lookup:
            cp_lookup[cp_code] += final_effective_value
        else:
            cp_lookup[cp_code] = final_effective_value

    # Round final totals
    # cp_lookup = {cp: round(val, 2) for cp, val in cp_lookup.items()}
    return cp_lookup

def read_file(file_path: str, header_row: int = 0, usecols=None) -> pd.DataFrame:
    """
    Dynamically read CSV, XLS, or XLSX file into a Pandas DataFrame.
    
    Parameters:
        file_path (str): Path to the input file.
        header_row (int): Row number (0-based) to use as header.
                          Example: header_row=2 means 3rd row is header.
                          
    Returns:
        pd.DataFrame: DataFrame containing file data.
    """
    # Ensure file exists
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    # Get file extension
    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == ".csv":
            try:
                df = pd.read_csv(
                    file_path,
                    header=header_row,
                    usecols=usecols,
                )
            except Exception:
                # fallback for inconsistent columns
                df = pd.read_csv(
                    file_path,
                    header=header_row,
                    usecols=usecols,
                    engine="python",
                    on_bad_lines="skip"  # skip problematic rows
                )
        elif ext == ".xlsx":
            df = pd.read_excel(file_path, header=header_row, usecols=usecols, engine="openpyxl")
        elif ext == ".xls":
            df = pd.read_excel(file_path, header=header_row, usecols=usecols, engine="xlrd")
        else:
            raise ValueError(f"Unsupported file type: {ext}")
    except Exception as e:
        raise RuntimeError(f"Error reading {file_path}: {str(e)}")
    
    return df

def write_file(file_path: str, data: list[dict], header: list[str]) -> None:
    """
    Write data into CSV, XLS, or XLSX with a given header order.
    
    Parameters:
        file_path (str): Output file path (.csv, .xls, .xlsx).
        data (list[dict]): List of dictionaries (rows).
        header (list[str]): Column names in required order.
        
    Returns:
        None
    """
    # Convert data to DataFrame with given header order
    df = pd.DataFrame(data)
    
    # Reorder columns (fill missing columns with empty values)
    for col in header:
        if col not in df.columns:
            df[col] = None
    df = df[header]
    
    # Get file extension
    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == ".csv":
            df.to_csv(file_path, index=False)
        elif ext == ".xlsx":
            # xlsxwriter gives clean output with styles, no warnings
            df.to_excel(file_path, index=False, engine="xlsxwriter")
        elif ext == ".xls":
            df.to_excel(file_path, index=False, engine="xlwt")  # legacy
        else:
            raise ValueError(f"Unsupported file type: {ext}")
    except Exception as e:
        raise RuntimeError(f"Error writing {file_path}: {str(e)}")

CASHCOLLATERAL_CDS = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\CashCollateral_cds.xls'
CASHCOLLATERAL_FNO = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\CashCollateral_fno.xls'
COLLATERAL_VIOLATION_REPORT = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\Collateral Valuation Report.xls'
DAILY_MARGIN_NSECR_FILE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\Daily Margin Report NSECR.xls'
DAILY_MARGIN_NSEFNO_FILE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\Daily Margin Report NSEFNO.xls'
FO_MSATER_FILE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\F_CPMaster_data.xlsx'
CD_MASTER_FILE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\X_CPMaster_data.xlsx'
SEC_PLEDGE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation\F_90123_SEC_PLEDGE_11092025_02.csv\F_90123_SEC_PLEDGE_11092025_02.csv'

# FO_MSATER_FILE = r"C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\F_CPMaster_data.xlsx"
# CD_MASTER_FILE = r"C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\X_CPMaster_data.xlsx"

# CASHCOLLATERAL_FNO = r"C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\CashCollateral_FNO.xls"  #
# CASHCOLLATERAL_CDS = r"C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\CashCollateral_CDS.xls"  # 

# DAILY_MARGIN_NSECR_FILE =  r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Daily Margin Report NSECR.xls'
# DAILY_MARGIN_NSEFNO_FILE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Daily Margin Report NSEFNO.xls'

# COLLATERAL_VIOLATION_REPORT = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Collateral Valuation Report.xls'
# SEC_PLEDGE = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\F_90123_SEC_PLEDGE_09092025_02.csv\F_90123_SEC_PLEDGE_09092025_02.csv'


date = "11-09-2025"
pan = "AACCO4820B"
account_type = "C"

# Variables
date = "11-09-2025"
pan = "AACCO4820B"
account_type = "C"

# Directory path
directory = r'C:\Users\KrishnaPatil\OneDrive - Dovetail Capital Pvt ltd\Desktop\PCM Files\BOD\Segregation'
# Filename using format
filename = "{}_{}.xlsx".format(date, pan)
outpath = os.path.join(directory, filename)

# File name without extension
FO = 'FO' if 'F' == os.path.splitext(os.path.basename(FO_MSATER_FILE))[0].split('_')[0] else ''
CD = 'CD' if os.path.splitext(os.path.basename(CD_MASTER_FILE))[0].split('_')[0] else ''

# Read FO file
df1 = read_file(FO_MSATER_FILE)
cp_codes_fo = df1["CP Code"].tolist()         # replace "CP Code" with your actual column name
pan_fo = df1["PAN Number"].tolist()           # replace "PAN Number" with actual column name

# Read CD file
df2 = read_file(CD_MASTER_FILE)
cp_codes_cd = df2["CP Code"].tolist()
pan_cd = df2["PAN Number"].tolist()

df3 = read_file(CASHCOLLATERAL_FNO, header_row=9, usecols="B:I")  # because row 10 is the header
fo_collateral_lookup = dict(zip(df3["ClientCode"], df3["TotalCollateral"]))

df4 = read_file(CASHCOLLATERAL_CDS, header_row=9, usecols="B:I")  # because row 10 is the header
cd_collateral_lookup = dict(zip(df4["ClientCode"], df4["TotalCollateral"]))

# "Funds" 
# "ClientCode"
df5 = read_file(DAILY_MARGIN_NSECR_FILE, header_row=9, usecols="B:T")  # because row 10 is the header
cd_daily_margin_lookup = dict(zip(df5["ClientCode"], df5["Funds"]))

df6 = read_file(DAILY_MARGIN_NSEFNO_FILE, header_row=9, usecols="B:T")  # because row 10 is the header
fo_daily_margin_lookup = dict(zip(df6["ClientCode"], df6["Funds"]))

df7 = read_file(COLLATERAL_VIOLATION_REPORT, header_row=9, usecols="B:H")
# Build lookup: {ClientCode: {"CashEquivalent": x, "NonCash": y}}
collateral_violation_lookup = {}

for _, row in df7.iterrows():
    client_code = row["ClientCode"]
    cash_eq = row["CashEquivalent"]
    non_cash = row["NonCash"]

    # If duplicate ClientCode found → handle manually
    if client_code in collateral_violation_lookup:
        # Example: take last value (overwrite)
        # collateral_violation_lookup[client_code] = {"CashEquivalent": cash_eq, "NonCash": non_cash}

        # OR: sum values instead
        collateral_violation_lookup[client_code]["CashEquivalent"] = cash_eq
        collateral_violation_lookup[client_code]["NonCash"] = non_cash
    else:
        collateral_violation_lookup[client_code] = {
            "CashEquivalent": cash_eq,
            "NonCash": non_cash
        }

# Step 1: Read raw file without header
header_row = None

# Step 1: Scan file for "GSEC" in first column
with open(SEC_PLEDGE, "r", encoding="utf-8", errors="ignore") as f:
    for idx, line in enumerate(f):
        first_col = line.split(",")[0].strip()  # only check first column
        if first_col.upper() == "GSEC":
            print(f"✅ Found 'GSEC' at line {idx}")
            header_row = idx + 1  # header is next line
            break

if header_row is None:
    raise ValueError("'GSEC' not found in first column of file!")

# Step 2: Read CSV using the detected header row
df8 = pd.read_csv(SEC_PLEDGE, header=header_row, engine="python")

# Step 5: Automatically strip column names
df8.columns = df8.columns.str.strip()

_sec_pledge_lookup = {}

for _, row in df8.iterrows():
    client_code = row['Client/CP code']
    isin = row['ISIN']
    gross_value = row['GROSS VALUE']
    haircut = row['HAIRCUT']

    if not client_code or not isin:
        continue  # skip incomplete rows

    key = f"{client_code}-{isin}"

    _sec_pledge_lookup[key] = {
        "GROSS VALUE": gross_value,
        "HAIRCUT": haircut
    }

sec_pledge_lookup = build_cp_lookup(_sec_pledge_lookup)

print("######## sec pledge lookup :", sec_pledge_lookup)

# Create list of lists
data = []


for cp, pan_no in zip(cp_codes_fo, pan_fo):
    cv_lookup = collateral_violation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
    row = {
        A : date,
        B : pan,
        C : pan,
        D : cp,
        E : pan_no,
        G : account_type,
        H : FO,
        J : fo_collateral_lookup.get(cp, 0),  # Default to 0 if CP Code not found
        K : fo_daily_margin_lookup.get(cp, 0),
        L : fo_daily_margin_lookup.get(cp, 0),
        O : cv_lookup["CashEquivalent"],
        P : cv_lookup["NonCash"],
        BB: sec_pledge_lookup.get(cp, 0),
        BD: sec_pledge_lookup.get(cp, 0),
        BF: sec_pledge_lookup.get(cp, 0)
    }
    
    # duplicate values in other columns
    row[AD] = row[K]
    row[AV] = row[K]
    row[AG] = row[O]
    row[AW] = row[O]
    row[AH] = row[P]
    row[AX] = row[P]

    data.append(row)

for cp, pan_no in zip(cp_codes_cd, pan_cd):
    cv_lookup = collateral_violation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
    row = {
        A : date,
        B : pan,
        C : pan,
        D : cp,
        E : pan_no,
        G : account_type,
        H : CD,
        J : cd_collateral_lookup.get(cp, 0),  # Default to 0 if CP Code not found
        K : cd_daily_margin_lookup.get(cp, 0),
        L : cd_daily_margin_lookup.get(cp, 0),
        O : cv_lookup["CashEquivalent"],
        P : cv_lookup["NonCash"],
    }
    # duplicate values in other columns
    row[AD] = row[K]
    row[AV] = row[K]
    row[AG] = row[O]
    row[AW] = row[O]
    row[AH] = row[P]
    row[AX] = row[P]

    data.append(row)

write_file(outpath, data=data, header=segregation_headers)

# CSV file, header is 3rd row
# df1 = read_file(FO_MSATER_FILE)

# # XLSX file, header is 2nd row
# df2 = read_file("report.xlsx", header_row=1)

# # XLS file, header is 1st row
# df3 = read_file("old_file.xls", header_row=0)

# this is for header
# print(df1.columns.tolist())
# # print(df1.info())
# print(df1.describe())