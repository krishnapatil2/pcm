"""
Report Processing Modules
Separated processing logic for different report types
"""

import csv
import os
import pandas as pd
import zipfile
import io
import calendar
from datetime import datetime, timedelta
import traceback
import glob
import cons_header
from db_manager import insert_report
from physical_settlement_files import build_dict, segregate_excel_by_column
from itertools import groupby
from operator import itemgetter
import json
import os

class BaseProcessor:
    """Base class for all processors"""
    def __init__(self, db_path, log_error_callback):
        self.db_path = db_path
        self.log_error = log_error_callback
    
    def validate_inputs(self, **kwargs):
        """Validate common inputs - override in subclasses"""
        pass
    
    def create_output_directory(self, output_path):
        """Create output directory if it doesn't exist"""
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path)
            except Exception as e:
                raise Exception(f"Cannot create output directory: {str(e)}")


class MonthlyFloatProcessor(BaseProcessor):
    """Processor for Monthly Float Report"""
    
    def process(self, fno_path, mcx_path, output_path):
        """Process FNO and MCX files for monthly float report"""
        try:
            self.validate_inputs(fno_path=fno_path, mcx_path=mcx_path, output_path=output_path)
            self.create_output_directory(output_path)
            
            error_log_path = os.path.join(output_path, "pcm_errors.txt")
            fno_files = glob.glob(os.path.join(fno_path, "*.csv"))
            mcx_files = glob.glob(os.path.join(mcx_path, "*.csv"))

            fno_count = len(fno_files)
            mcx_count = len(mcx_files)
            
            # Merge files
            self._merge_fno_and_mcx(fno_files, mcx_files, output_path, error_log_path)
            
            # Process data
            df_list = []
            for folder in [fno_files, mcx_files]:
                for file in folder:
                    try:
                        temp_df = pd.read_csv(file)
                        temp_df.columns = temp_df.columns.str.strip()
                        df_list.append(temp_df[cons_header.columns_to_keep])
                    except Exception as e:
                        self.log_error(error_log_path, file, e)
                        print(f"Error reading {file}: {e}")

            if not df_list:
                raise Exception("No CSV files found or all failed to load.")

            df = pd.concat(df_list, ignore_index=True)

            # Fill missing dates
            df_before_fill = len(df)
            df, messages = self._fill_missing_dates(df, error_log_path)
            df_after_fill = len(df)
            missing_filled = df_after_fill - df_before_fill

            # Generate summary
            summary_data = self._generate_summary(df)
            
            # Write Excel file
            output_file = self._write_excel_output(df, summary_data, output_path)
            
            # Create monthly status log
            self._create_monthly_status_log(messages, output_path)
            
            # Create ZIP and save to database
            self._create_zip_and_save(output_path, fno_path, mcx_path, output_file)
            
            return {
                'fno_count': fno_count,
                'mcx_count': mcx_count,
                'missing_filled': missing_filled,
                'output_file': output_file
            }
            
        except Exception as e:
            self.log_error(output_path, "Monthly Float Processing", e)
            raise e
    
    def validate_inputs(self, fno_path, mcx_path, output_path):
        """Validate inputs for monthly float processing"""
        if not fno_path or not mcx_path or not output_path:
            raise ValueError("Please select all folders before processing.")
    
    def _merge_fno_and_mcx(self, fno_files, mcx_files, output_path, error_log_path):
        """Merge FNO and MCX data"""
        try:
            all_dataframes = []

            for file in fno_files:
                try:
                    df = pd.read_csv(file)
                    df['Source'] = 'FNO'
                    all_dataframes.append(df)
                except Exception as e:
                    self.log_error(error_log_path, file, e)

            for file in mcx_files:
                try:
                    df = pd.read_csv(file)
                    df['Source'] = 'MCX'
                    all_dataframes.append(df)
                except Exception as e:
                    self.log_error(error_log_path, file, e)

            if not all_dataframes:
                raise Exception(f"No valid CSV files found in FNO or MCX.\nSee log: {error_log_path}")

            merged_df = pd.concat(all_dataframes, ignore_index=True)
            output_file = os.path.join(output_path, "merged_fno_mcx_data.xlsx")
            merged_df.to_excel(output_file, index=False)

        except Exception as e:
            self.log_error(error_log_path, "merge_fno_and_mcx", e)
            raise Exception(f"A fatal error occurred.\nSee log: {error_log_path}")
    
    def _fill_missing_dates(self, df, error_log_path):
        """Fill missing dates by duplicating previous day's data"""
        try:
            df[cons_header.DATE] = pd.to_datetime(df[cons_header.DATE], dayfirst=True)
            df["YEAR"] = df[cons_header.DATE].dt.year
            df["MONTH"] = df[cons_header.DATE].dt.month

            filled_data = []
            status_messages = []

            for cp_code, cp_group in df.groupby(cons_header.CP_CODE, dropna=False):
                for (year, month), month_group in cp_group.groupby(["YEAR", "MONTH"]):
                    try:
                        year = int(year)
                        month = int(month)

                        days_in_month = calendar.monthrange(year, month)[1]
                        month_name = calendar.month_name[month]

                        all_days = pd.date_range(start=f"{year}-{month:02d}-01", periods=days_in_month)

                        existing_dates = set(month_group[cons_header.DATE])
                        missing_dates = [d for d in all_days if d not in existing_dates]

                        msg = ""
                        cp_code_display = "blankcpcode" if cp_code == "" or str(cp_code).lower() == "nan" else cp_code
                        if missing_dates:
                            missing_day_nums = ", ".join(str(d.day) for d in missing_dates)
                            msg = f"[INFO] {cp_code_display} → {month_name} {year}: Missing {len(missing_dates)} day(s) → Days: {missing_day_nums}"
                        else:
                            msg = f"[SUCCESS] {cp_code_display} → {month_name} {year}: ✅ All {days_in_month} days present."

                        status_messages.append(msg)
                        filled_month = month_group.copy()

                        for date in missing_dates:
                            prev_data = filled_month[filled_month[cons_header.DATE] < date]
                            if prev_data.empty:
                                continue
                            last_row = prev_data.sort_values(cons_header.DATE).iloc[-1].copy()
                            last_row[cons_header.DATE] = date
                            filled_month = pd.concat([filled_month, pd.DataFrame([last_row])])

                        filled_data.append(filled_month.sort_values(cons_header.DATE))

                    except Exception as e:
                        if error_log_path:
                            self.log_error(error_log_path, f"{cp_code} - {year}-{month:02d}", e)
                        return None, []

            final_df = pd.concat(filled_data, ignore_index=True)
            final_df.drop(columns=["YEAR", "MONTH"], inplace=True)
            return final_df, status_messages

        except Exception as e:
            if error_log_path:
                self.log_error(error_log_path, "fill_missing_dates", e)
            raise e
    
    def _generate_summary(self, df):
        """Generate summary data for the report"""
        summary_data = []
        df[cons_header.CP_CODE] = df[cons_header.CP_CODE].fillna("").astype(str).replace("nan", "").str.strip()
        cp_groups = list(df.groupby(cons_header.CP_CODE, dropna=False))
        
        for cp_code, group in cp_groups:
            cp_code_display = "blankcpcode" if cp_code == "" else cp_code
            total_val = group[cons_header.FINANCIAL_LEDGER_BALANCE].sum()
            avg_val = group[cons_header.FINANCIAL_LEDGER_BALANCE].mean()
            group[cons_header.CP_CODE] = "" if cp_code == "" else cp_code
            summary_data.append({
                "CP Code": cp_code_display,
                "Total": total_val,
                "Average": avg_val
            })
        
        return summary_data
    
    def _write_excel_output(self, df, summary_data, output_path):
        """Write Excel output file"""
        output_file = os.path.join(output_path, "cp_code_separate_sheets.xlsx")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            summary_df = pd.DataFrame(summary_data, columns=["CP Code", "Total", "Average"])
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            df[cons_header.CP_CODE] = df[cons_header.CP_CODE].fillna("").astype(str).replace("nan", "").str.strip()
            cp_groups = list(df.groupby(cons_header.CP_CODE, dropna=False))
            
            for cp_code, group in cp_groups:
                sheet_name = "blankcpcode" if cp_code == "" else cp_code[:31]
                group[cons_header.CP_CODE] = "" if cp_code == "" else cp_code
                group.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return output_file
    
    def _create_monthly_status_log(self, messages, output_path):
        """Create monthly status log file"""
        monthly_log_path = os.path.join(output_path, "monthly_status.txt")
        user_friendly_header = "ℹ️ Monthly Status: Missing dates have been filled automatically. Please check the summary below.\n\n"
        full_message = user_friendly_header + "\n".join(messages)

        with open(monthly_log_path, "w", encoding="utf-8") as f:
            f.write(full_message)
    
    def _create_zip_and_save(self, output_path, fno_path, mcx_path, output_file):
        """Create ZIP file and save to database"""
        # Create ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(output_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_path)
                    zipf.write(file_path, arcname)
        zip_blob = zip_buffer.getvalue()

        # Insert into database
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        insert_report(self.db_path, report_type=cons_header.NSE_AND_MCX, 
                     created_at=timestamp, modified_at=timestamp, report_blob=zip_blob)


class NMASSAllocationProcessor(BaseProcessor):
    """Processor for NMASS Allocation Report"""
    
    def process(self, date, sheet, input1_path, input2_path, output_path):
        """Process NMASS allocation files"""
        try:
            self.validate_inputs(date=date, input1_path=input1_path, 
                               input2_path=input2_path, output_path=output_path)
            self.create_output_directory(output_path)
            
            # Process the files
            result = self._process_ledger_files(input1_path, input2_path, date, sheet, output_path)
            
            if result:
                return f"Ledger processed successfully with {result['record_count']} records."
            else:
                raise Exception("Failed to process ledger files.")
                
        except Exception as e:
            self.log_error(output_path, "NMASS Allocation Processing", e)
            raise e
    
    def validate_inputs(self, date, input1_path, input2_path, output_path):
        """Validate inputs for NMASS allocation processing"""
        if date == "DD/MM/YYYY" or not date.strip():
            raise ValueError("Please select a valid date.")
            
        if not input1_path.strip() or not input2_path.strip():
            raise ValueError("Please select both attachment files.")
        
        if not output_path.strip():
            raise ValueError("Please select an output folder for the ledger.")
        
        if not os.path.exists(input1_path):
            raise ValueError(f"Attachment 1 file not found:\n{input1_path}")
            
        if not os.path.exists(input2_path):
            raise ValueError(f"Attachment 2 file not found:\n{input2_path}")
    
    def _read_file(self, file_path, selected_sheet=None, **kwargs):
        """Read file with appropriate method based on extension"""
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == ".csv":
            df = pd.read_csv(file_path, **kwargs)
        elif ext in [".xls", ".xlsx"]:
            df = pd.read_excel(file_path, sheet_name=selected_sheet or 0, **kwargs)
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
        # Drop rows where all columns are NaN
        df = df.dropna(how='all')
        return df
    
    def _get_next_file_path(self, output_path, base_name, dt):
        """Generate the next available file path by incrementing T000X"""
        import re
        pattern = re.compile(rf"{re.escape(base_name)}_ALLOC_{dt}\.T(\d+)$")

        max_num = 0
        for fname in os.listdir(output_path):
            match = pattern.match(fname)
            if match:
                num = int(match.group(1))
                if num > max_num:
                    max_num = num

        next_num = max_num + 1
        filename = f"{base_name}_ALLOC_{dt}.T{next_num:04d}"
        return os.path.normpath(os.path.join(output_path, filename))
    
    def _build_segment_line(self, date, segment, member_code, cp_code, c_value, margin_value, status):
        """Build segment line for the report"""
        return f"{date},{segment},{member_code},,{cp_code},,{c_value},{margin_value},,,,,,,{status}"
    
    def _process_ledger_files(self, file1_path, file2_path, selected_date, selected_sheet, output_path):
        """Process ledger files and perform calculations"""
        try:
            df1 = self._read_file(file1_path)
            ext = os.path.splitext(file2_path)[1].lower()

            if ext == ".csv":
                df2 = self._read_file(file2_path, header=9, usecols="B:K")
            elif ext in [".xls", ".xlsx"]:
                df2 = self._read_file(file2_path, header=9, usecols=[cons_header.CLIENT_CODE, cons_header.FO_MARGIN])
            else:
                raise ValueError(f"Unsupported file type: {ext}")
            
            df1[cons_header.TM_CP_CODE] = df1[cons_header.TM_CP_CODE].astype(str)
            df2[cons_header.CLIENT_CODE] = df2[cons_header.CLIENT_CODE].astype(str)

            data_dict = {
                "file_1": dict(zip(df1[cons_header.TM_CP_CODE], df1[cons_header.CASH_COLLECTED])),
                "file_2": dict(zip(df2[cons_header.CLIENT_CODE], df2[cons_header.FO_MARGIN]))
            }
            
            dict1 = data_dict["file_1"]
            dict2 = data_dict["file_2"]

            formatted_date = datetime.strptime(selected_date, "%d/%m/%Y").strftime("%d-%b-%Y").upper()
            dt = datetime.strptime(selected_date, "%d/%m/%Y").strftime("%d%m%Y")
            processed_lines = set()

            file_path = self._get_next_file_path(output_path, cons_header.NSE_MEMBER_CODE, dt)
            lines_to_write = []

            # Process existing keys
            for key in dict1:
                if key.lower() == "nan":
                    continue
                if key in dict2:
                    difference = dict2[key] - dict1[key]
                    
                    if difference > 0:
                        status = "U"
                    elif difference < 0:
                        status = "D"
                    else:
                        continue

                    line_fo = self._build_segment_line(formatted_date, cons_header.SEGMENTS[selected_sheet], 
                                                     cons_header.NSE_MEMBER_CODE, key, cons_header.C, dict2[key], status)

                    if line_fo not in processed_lines:
                        lines_to_write.append(line_fo)
                        processed_lines.add(line_fo)
            
            # Process keys in dict2 but not in dict1
            for key in dict2:
                if key.lower() == "nan":
                    continue
                if key not in dict1:
                    if float(dict2[key]) == 0:
                        continue
                    status = "U"
                    line_fo = self._build_segment_line(formatted_date, cons_header.SEGMENTS[selected_sheet], 
                                                     cons_header.NSE_MEMBER_CODE, key, cons_header.C, dict2[key], status)
                    if line_fo not in processed_lines:
                        lines_to_write.append(line_fo)
                        processed_lines.add(line_fo)
            
            # Sort so that 'D' comes before 'U'
            sorted_lines = sorted(lines_to_write, key=lambda x: x.strip().split(",")[-1])
            for i in sorted_lines:
                if i.split(",")[4] == '90072':
                    sorted_lines.remove(i)

            # Write lines into report file
            with open(file_path, "w") as f:
                if lines_to_write:
                    f.write("\n".join(sorted_lines))
                else:
                    f.write("")

            # Create ZIP and save to database
            self._create_zip_and_save(file1_path, file2_path, file_path, output_path, dt)
            
            return {"record_count": len(sorted_lines)}

        except Exception as e:
            print(f"Error in process_ledger_files: {e}")
            self.log_error(output_path, "Error in process_ledger_files", e)
            return None
    
    def _create_zip_and_save(self, file1_path, file2_path, output_file, output_path, dt):
        """Create ZIP file and save to database"""
        # Create ZIP
        zip_filename = f"{cons_header.NSE_MEMBER_CODE}_REPORT_{dt}.zip"
        zip_path = os.path.join(output_path, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(file1_path, os.path.basename(file1_path))
            zipf.write(file2_path, os.path.basename(file2_path))
            zipf.write(output_file, os.path.basename(output_file))

        # Read ZIP as binary and insert into DB
        with open(zip_path, 'rb') as f:
            zip_blob = f.read()

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        insert_report(self.db_path, report_type=cons_header.LEDGER, 
                     created_at=timestamp, modified_at=timestamp, report_blob=zip_blob)


class ObligationSettlementProcessor(BaseProcessor):
    """Processor for Obligation Settlement"""
    
    def process(self, obligation_path, stt_path, stamp_duty_path, output_path):
        """Process obligation settlement files"""
        try:
            self.validate_inputs(obligation_path=obligation_path, stt_path=stt_path, 
                               stamp_duty_path=stamp_duty_path, output_path=output_path)
            self.create_output_directory(output_path)
            
            # Build dictionaries
            stt_dict = build_dict(
                file_path=stt_path,
                key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
                value_cols={
                    "Buy STT": "BuyDelvryTtlTaxs",
                    "Sell STT": "SellDelvryTtlTaxs"
                },
                filter_col="RptHdr",
                filter_val=40
            )

            stamp_duty_dict = build_dict(file_path=stamp_duty_path,
                key_cols=["BrkrOrCtdnPtcptId","TckrSymb", "FinInstrmId"],
                value_cols={
                    "Buy Stamp Duty": "BuyDlvryStmpDty"
                },
                filter_col="RptHdr",
                filter_val=40
            )
            
            output_file = os.path.join(output_path, "Physical_Settlement_Report.xlsx")

            # Segregate obligation and update Buy/Sell STT from dictionary
            segregate_excel_by_column(
                excel_path=obligation_path,
                output_path=output_file,
                column_name="BrkrOrCtdnPtcptId",
                custom_header=cons_header.OBLIGATION_HEADER,
                update_dicts=[stt_dict, stamp_duty_dict]
            )

            # Create ZIP and save to database
            self._create_zip_and_save(obligation_path, stt_path, stamp_duty_path, output_file, output_path)
            
            return f"Physical Settlement processed successfully. Output: {output_file}"

        except Exception as e:
            self.log_error(output_path, "Physical Settlement Processing", e)
            raise e
    
    def validate_inputs(self, obligation_path, stt_path, stamp_duty_path, output_path):
        """Validate inputs for obligation settlement processing"""
        if not obligation_path or not stt_path or not stamp_duty_path or not output_path:
            raise ValueError("Please select all input files and output folder.")
    
    def _create_zip_and_save(self, obligation_path, stt_path, stamp_duty_path, output_file, output_path):
        """Create ZIP file and save to database"""
        # Create ZIP
        dt = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"{cons_header.NSE_MEMBER_CODE}_PHYSICAL_SETTLEMENT_{dt}.zip"
        zip_path = os.path.join(output_path, zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(obligation_path, os.path.basename(obligation_path))
            zipf.write(stt_path, os.path.basename(stt_path))
            zipf.write(stamp_duty_path, os.path.basename(stamp_duty_path))
            zipf.write(output_file, os.path.basename(output_file))

        # Insert ZIP into DB
        with open(zip_path, 'rb') as f:
            zip_blob = f.read()
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        insert_report(self.db_path, report_type="PHYSICAL_SETTLEMENT", 
                     created_at=timestamp, modified_at=timestamp, report_blob=zip_blob)


class SegregationReportProcessor(BaseProcessor):
    """Processor for Segregation Report"""
    
    def process(self, date, cp_pan, cash_collateral_cds, cash_collateral_fno,
                daily_margin_nsecr, daily_margin_nsefno, x_cp_master, f_cp_master,
                collateral_valuation_cds, collateral_valuation_fno,
                # sec_pledge, 
                cash_with_ncl, santom_file, extra_records, output_path):
        """Process segregation report files"""
        try:
            self.validate_inputs(date=date, cp_pan=cp_pan, output_path=output_path,
                               cash_collateral_cds=cash_collateral_cds, cash_collateral_fno=cash_collateral_fno,
                               daily_margin_nsecr=daily_margin_nsecr, daily_margin_nsefno=daily_margin_nsefno,
                               x_cp_master=x_cp_master, f_cp_master=f_cp_master,
                               collateral_valuation_cds=collateral_valuation_cds, collateral_valuation_fno=collateral_valuation_fno, 
                            #   sec_pledge=sec_pledge,
                               cash_with_ncl=cash_with_ncl, santom_file=santom_file, extra_records=extra_records)
            self.create_output_directory(output_path)
            
            # Process the segregation report
            result = self._process_segregation_files(
                date, cp_pan, cash_collateral_cds, cash_collateral_fno,
                daily_margin_nsecr, daily_margin_nsefno, x_cp_master, f_cp_master,
                collateral_valuation_cds, collateral_valuation_fno, 
                # sec_pledge, 
                cash_with_ncl, santom_file, extra_records, output_path
            )
            
            return result
                
        except Exception as e:
            self.log_error(output_path, "Segregation Report Processing", e)
            raise e
    
    def validate_inputs(self, date, cp_pan, output_path, **file_paths):
        """Validate inputs for segregation report processing"""
        if date == "DD/MM/YYYY" or not date.strip():
            raise ValueError("Please select a valid date.")
            
        if not cp_pan.strip():
            raise ValueError("Please enter CP PAN.")
        
        # Check required files
        missing_files = []
        for file_name, file_path in file_paths.items():
            if not file_path.strip():
                missing_files.append(file_name.replace('_', ' ').title())
        
        if missing_files:
            raise ValueError(f"Please select the following files:\n" + "\n".join(f"- {name}" for name in missing_files))
        
        if not output_path.strip():
            raise ValueError("Please select an output folder.")
        
        # Check if files exist
        for file_name, file_path in file_paths.items():
            
            if file_name == "cash_with_ncl":
                continue

            if file_path and not os.path.exists(file_path):
                raise ValueError(f"File not found: {file_name.replace('_', ' ').title()}\n{file_path}")
    
    def _process_segregation_files(self, date, cp_pan, cash_collateral_cds, cash_collateral_fno,
                                daily_margin_nsecr, daily_margin_nsefno, x_cp_master, f_cp_master,
                                collateral_valuation_cds, collateral_valuation_fno, 
                                # sec_pledge, 
                                cash_with_ncl, santom_file, extra_records, output_path):
        """Process all segregation files and generate the final report"""
        try:
            # Import segregation functions
            from segregation import read_file, write_file, build_cp_lookup
            from CONSTANT_SEGREGATION import segregation_headers, A, B, C, D, E, F, G, H, I, J, K, L, O, P, AD, AV, AG, AW, AH, AX, BB, BD, BF, AT
            
            # Format date for output
            formatted_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d-%m-%Y")
            
            # Read CP Master files
            try:
                df_fo = read_file(f_cp_master)
                cp_codes_fo = df_fo["CP Code"].tolist()
                pan_fo = df_fo["PAN Number"].tolist()
            except Exception as e:
                raise Exception(f"❌ Error reading F_CPMaster_data file:\n\nPlease check if the correct F_CPMaster_data file is attached.\n\nTechnical details: {str(e)}")
            
            try:
                df_cd = read_file(x_cp_master)
                cp_codes_cd = df_cd["CP Code"].tolist()
                pan_cd = df_cd["PAN Number"].tolist()
            except Exception as e:
                raise Exception(f"❌ Error reading X_CPMaster_data file:\n\nPlease check if the correct X_CPMaster_data file is attached.\n\nTechnical details: {str(e)}")
            
            # Read Cash Collateral files
            try:
                df_cash_cds = read_file(cash_collateral_cds, header_row=9, usecols="B:I")
                cd_collateral_lookup = dict(zip(df_cash_cds["ClientCode"], df_cash_cds["TotalCollateral"]))
            except Exception as e:
                raise Exception(f"❌ Error reading CashCollateral_CDS file:\n\nPlease check if the correct CashCollateral_CDS file is attached.\n\nTechnical details: {str(e)}")
            
            try:
                df_cash_fno = read_file(cash_collateral_fno, header_row=9, usecols="B:I")
                fo_collateral_lookup = dict(zip(df_cash_fno["ClientCode"], df_cash_fno["TotalCollateral"]))
            except Exception as e:
                raise Exception(f"❌ Error reading CashCollateral_FNO file:\n\nPlease check if the correct CashCollateral_FNO file is attached.\n\nTechnical details: {str(e)}")
            
            # Read Daily Margin files
            try:
                df_margin_cds = read_file(daily_margin_nsecr, header_row=9, usecols="B:T")
                cd_daily_margin_lookup = dict(zip(df_margin_cds["ClientCode"], df_margin_cds["Funds"]))
            except Exception as e:
                raise Exception(f"❌ Error reading Daily Margin Report NSECR file:\n\nPlease check if the correct Daily Margin Report NSECR file is attached.\n\nTechnical details: {str(e)}")
            
            try:
                df_margin_fno = read_file(daily_margin_nsefno, header_row=9, usecols="B:T")
                fo_daily_margin_lookup = dict(zip(df_margin_fno["ClientCode"], df_margin_fno["Funds"]))
            except Exception as e:
                raise Exception(f"❌ Error reading Daily Margin Report NSEFNO file:\n\nPlease check if the correct Daily Margin Report NSEFNO file is attached.\n\nTechnical details: {str(e)}")
            
            # Read Collateral Violation Report CD
            try:
                df_valuation_cd = read_file(collateral_valuation_cds, header_row=9, usecols="B:H")
                cd_collateral_valuation_lookup = {}

                for _, row in df_valuation_cd.iterrows():
                    client_code = row["ClientCode"]
                    cash_eq = row["CashEquivalent"]
                    non_cash = row["NonCash"]
                    
                    if client_code in cd_collateral_valuation_lookup:
                        cd_collateral_valuation_lookup[client_code]["CashEquivalent"] = cash_eq
                        cd_collateral_valuation_lookup[client_code]["NonCash"] = non_cash
                    else:
                        cd_collateral_valuation_lookup[client_code] = {
                            "CashEquivalent": cash_eq,
                            "NonCash": non_cash
                        }
            except Exception as e:
                raise Exception(f"❌ Error reading Collateral Violation Report CDS file:\n\nPlease check if the correct Collateral Violation Report CDS file is attached.\n\nTechnical details: {str(e)}")

            # Read Collateral Violation Report FO
            try:
                df_valuation_fo = read_file(collateral_valuation_fno, header_row=9, usecols="B:H")
                fo_collateral_valuation_lookup = {}
                
                for _, row in df_valuation_fo.iterrows():
                    client_code = row["ClientCode"]
                    cash_eq = row["CashEquivalent"]
                    non_cash = row["NonCash"]
                    
                    if client_code in fo_collateral_valuation_lookup:
                        fo_collateral_valuation_lookup[client_code]["CashEquivalent"] = cash_eq
                        fo_collateral_valuation_lookup[client_code]["NonCash"] = non_cash
                    else:
                        fo_collateral_valuation_lookup[client_code] = {
                            "CashEquivalent": cash_eq,
                            "NonCash": non_cash
                        }
            except Exception as e:
                raise Exception(f"❌ Error reading Collateral Violation Report FNO file:\n\nPlease check if the correct Collateral Violation Report FNO file is attached.\n\nTechnical details: {str(e)}")
            
            # Process Security Pledge file
            sec_pledge_cp_lookup = {} #self._process_security_pledge_file(sec_pledge)
            
            # Generate report data
            data = self._generate_report_data(
                formatted_date, cp_pan, cp_codes_fo, pan_fo, cp_codes_cd, pan_cd,
                fo_collateral_lookup, fo_daily_margin_lookup, cd_collateral_lookup, 
                cd_daily_margin_lookup, cd_collateral_valuation_lookup,fo_collateral_valuation_lookup, sec_pledge_cp_lookup
            )
            data = self._segregation_data_filter(data, segregation_headers=segregation_headers[9:])
            # breakpoint()

            # Load master records using simple dynamic function
            av_records, at_records = self._get_master_records() # Get Both AV and AT Records (Default):
            # 2. Get Only AV or AT Records:
            # av_records = self._get_master_records(av=True) at_records = self._get_master_records(at=True)
            # all_records = self._get_master_records(all_records=True)
            
            # Add extra records
            if extra_records:
                try:
                    extra_records_df = read_file(extra_records)
                    for _, row in extra_records_df.iterrows():
                        record = {}
                        for col in extra_records_df.columns:
                            val = row[col]

                            if col == A:
                                try:
                                    if isinstance(val, pd.Timestamp):
                                        val = val.strftime("%d-%m-%Y")
                                    elif isinstance(val, str):
                                        # normalize strings like "2025-09-18 00:00:00"
                                        val = val.split(" ")[0]
                                    else:
                                        val = ""
                                except Exception:
                                    val = ""
                                
                            record[col] = val
                        
                        # Custom logic
                        # if str(row.get(G, "")).strip() == "P" and str(row.get(H, "")).strip() == "FO":
                        #     # Lookup in AV_Records
                        #     for av_record in av_records:
                        #         if (
                        #             av_record.get(G) == "P" and
                        #             av_record.get(H) == "FO"
                        #         ):
                        #             record[AV] = av_record["av_value"]
                        #             break  # stop at first match

                        data.append(record)
                except Exception as e:
                    raise Exception(f"❌ Error reading Extra_Records_File:\n\nPlease check if the correct Extra_Records_File is attached.\n\nTechnical details: {str(e)}")
            
            # Loop through data (list of dictionaries) and apply AT records logic
            # print(f" Processing data with AT records logic...")
            for i, data_record in enumerate(data):
                # Loop through AT records to find matches
                for at_record in at_records:
                    at_cp_code = at_record.get(D, '')
                    at_segment = at_record.get(H, '')
                    at_value = float(at_record.get("at_value", 0))
                    
                    # Check if current data record matches AT record criteria
                    if (data_record.get(D, '') == at_cp_code and 
                        data_record.get(H, '') == at_segment):
                        
                        # Apply AT logic to this data record

                        data_record[AV] = data_record[AD] - at_value
                        # print(f"  Applied AT logic to record {i+1}: CP Code={at_cp_code}, Segment={at_segment}, AT Value={at_value}")
                        break  # Stop at first match

            try:
                santom_df = read_file(santom_file)
                data = self._santom_file_working(data, cash_with_ncl, santom_df)
            except Exception as e:
                raise Exception(f"❌ Error reading SANTOM_FILE:\n\nPlease check if the correct SANTOM_FILE is attached.\n\nTechnical details: {str(e)}")

            # Write output file
            output_file = os.path.join(output_path, f"{cp_pan}_{formatted_date.replace('-', '')}_01.xlsx")
            write_file(output_file, data=data, header=segregation_headers)

            # Create ZIP and save to database
            self._create_zip_and_save(
                cash_collateral_cds, cash_collateral_fno, daily_margin_nsecr, daily_margin_nsefno,
                x_cp_master, f_cp_master, collateral_valuation_cds, collateral_valuation_fno, 
                # sec_pledge,
                output_file, output_path
            )
            
            return f"Segregation report generated successfully with {len(data)} records."
            
        except Exception as e:
            print(f"Error in process_segregation_files: {e}")
            self.log_error(output_path, "Error in process_segregation_files", e)
            return None
    
    def _process_security_pledge_file(self, sec_pledge):
        """Process security pledge file"""
        # Step 1: Scan file for "GSEC" in first column
        # header_row = None
        # with open(sec_pledge, "r", encoding="utf-8", errors="ignore") as f:
        #     for idx, line in enumerate(f):
        #         first_col = line.split(",")[0].strip()
        #         if first_col.upper() == "GSEC":
        #             print(f"✅ Found 'GSEC' at line {idx}")
        #             header_row = idx + 1
        #             break

        # if header_row is None:
        #     raise ValueError("'GSEC' not found in first column of file!")

        # # Step 2: Read CSV using the detected header row
        # df8 = pd.read_csv(sec_pledge, header=header_row, engine="python")
        # df8.columns = df8.columns.str.strip()

        # _sec_pledge_lookup = {}
        # for _, row in df8.iterrows():
        #     client_code = row['Client/CP code']
        #     isin = row['ISIN']
        #     gross_value = row['GROSS VALUE']
        #     haircut = row['HAIRCUT']

        #     if not client_code or not isin:
        #         continue

        #     key = f"{client_code}-{isin}"
        #     _sec_pledge_lookup[key] = {
        #         "GROSS VALUE": gross_value,
        #         "HAIRCUT": haircut
        #     }

        _sec_pledge_lookup = {}

        with open(sec_pledge, newline='', encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            rows = list(reader)

        # Step 1: Find where "GSEC" occurs in first column
        header_row = None
        for idx, row in enumerate(rows):
            if row and row[0].strip().upper() == "GSEC":
                print(f"✅ Found 'GSEC' at line {idx}")
                # Prefer next line if it looks like header
                header_row = idx + 1
                break

        if header_row is None:
            raise ValueError("'GSEC' not found in file")

        # Step 2: Extract headers and data
        headers = [col.strip() for col in rows[header_row]]
        data_rows = rows[header_row + 1:]

        # Step 3: Build lookup dictionary
        try:
            col_client = headers.index("Client/CP code")
            col_isin = headers.index("ISIN")
            col_gross = headers.index("GROSS VALUE")
            col_haircut = headers.index("HAIRCUT")
        except ValueError as e:
            raise ValueError(f"❌ Expected column missing: {e}")

        for row in data_rows:
            if len(row) <= max(col_client, col_isin, col_gross, col_haircut):
                continue  # skip short/incomplete rows

            client_code = row[col_client].strip()
            isin = row[col_isin].strip()
            gross_value = row[col_gross].strip()
            haircut = row[col_haircut].strip()

            if not client_code or not isin:
                continue

            key = f"{client_code}-{isin}"
            _sec_pledge_lookup[key] = {
                "GROSS VALUE": gross_value,
                "HAIRCUT": haircut,
            }

        from segregation import build_cp_lookup
        return build_cp_lookup(_sec_pledge_lookup)
    
    def _generate_report_data(self, formatted_date, cp_pan, cp_codes_fo, pan_fo, cp_codes_cd, pan_cd,
                            fo_collateral_lookup, fo_daily_margin_lookup, cd_collateral_lookup, 
                            cd_daily_margin_lookup, cd_collateral_valuation_lookup, fo_collateral_valuation_lookup, sec_pledge_cp_lookup):
        """Generate report data for both FO and CD segments"""
        from CONSTANT_SEGREGATION import A, B, C, D, E, F, G, H, I, J, K, L, O, P, AD, AV, AG, AW, AH, AX, BB, BD, BF
        
        data = []
        account_type = "C"
        
        # Process FO data
        for cp, pan_no in zip(cp_codes_fo, pan_fo):
            cv_lookup = fo_collateral_valuation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
            row = {
                A: formatted_date,
                B: cp_pan,
                C: cp_pan,
                D: cp,
                E: pan_no,
                F: "",  # Client PAN
                G: account_type,
                H: "FO",
                I: "",  # UCC Code
                J: fo_collateral_lookup.get(cp, 0),
                K: fo_daily_margin_lookup.get(cp, 0),
                L: fo_daily_margin_lookup.get(cp, 0),
                O: cv_lookup["CashEquivalent"],
                P: cv_lookup["NonCash"],
                BB: cv_lookup["CashEquivalent"],
                BD: cv_lookup["CashEquivalent"],
                BF: cv_lookup["CashEquivalent"]
            }
            
            # Duplicate values in other columns
            row[AD] = row[K]
            row[AV] = row[K]
            row[AG] = row[O]
            row[AW] = row[O]
            row[AH] = row[P]
            row[AX] = row[P]
            
            data.append(row)
        
        # Process CD data
        for cp, pan_no in zip(cp_codes_cd, pan_cd):
            cv_lookup = cd_collateral_valuation_lookup.get(cp, {"CashEquivalent": 0, "NonCash": 0})
            row = {
                A: formatted_date,
                B: cp_pan,
                C: cp_pan,
                D: cp,
                E: pan_no,
                F: "",  # Client PAN
                G: account_type,
                H: "CD",
                I: "",  # UCC Code
                J: cd_collateral_lookup.get(cp, 0),
                K: cd_daily_margin_lookup.get(cp, 0),
                L: cd_daily_margin_lookup.get(cp, 0),
                O: cv_lookup["CashEquivalent"],
                P: cv_lookup["NonCash"],
                BB: cv_lookup["CashEquivalent"],
                BD: cv_lookup["CashEquivalent"],
                BF: cv_lookup["CashEquivalent"]
            }
            
            # Duplicate values in other columns
            row[AD] = row[K]
            row[AV] = row[K]
            row[AG] = row[O]
            row[AW] = row[O]
            row[AH] = row[P]
            row[AX] = row[P]
            
            data.append(row)
        
        return data
    
    def _segregation_data_filter(self, data, segregation_headers, cp_code_col="CP Code", seg_col="Segment Indicator"):
        """
        Filter and normalize segregation data:
        1. Replace blank/NA values with 0 for segregation_headers
        2. Sort by CP Code
        3. Sort by Segment Indicator
        4. Move all-zero rows to the end

        Args:
            data (list[dict]): list of row dictionaries
            segregation_headers (list[str]): expected headers for segregation
            cp_code_col (str): CP Code column name
            seg_col (str): Segment Indicator column name
        
        Returns:
            list[dict]: filtered and sorted data
        """
        # Step 1: Normalize data - replace blank/NA values with 0 for segregation_headers
        normalized = []
        for row in data:
            new_row = {}
            
            # Process segregation_headers columns
            for col in segregation_headers:
                val = row.get(col, 0)  # default if missing
                if val is None or (isinstance(val, str) and (val.strip() == "" or val.strip().upper() == "NA")):
                    val = 0
                new_row[col] = val

            # Copy other columns as-is
            for key, val in row.items():
                if key not in segregation_headers:
                    new_row[key] = val

            normalized.append(new_row)

        # Step 2 & 3: Sort by CP Code A to Z, then by Segment Indicator A to Z within each CP Code group
        seg_sorted = sorted(normalized, key=lambda x: (str(x.get(cp_code_col, "")).strip().upper(), str(x.get(seg_col, "")).strip().upper()))
        
        # Step 4: Move all-zero rows to the end
        def is_all_zero(row):
            """Check if all segregation_headers columns have zero values"""
            return all(
                (v == 0 or v == "0" or str(v).strip() == "0") 
                for k, v in row.items() 
                if k in segregation_headers
            )

        # Separate zero and non-zero rows
        zero_rows = []
        non_zero_rows = []
        
        for row in seg_sorted:
            if is_all_zero(row):
                zero_rows.append(row)
            else:
                non_zero_rows.append(row)
        
        # Combine: non-zero rows first, then zero rows
        final_data = non_zero_rows + zero_rows

        return final_data

    def _santom_file_working(self, data, cash_with_ncl, santom_df):
        from CONSTANT_SEGREGATION import segregation_headers, A, B, C, D, E, F, G, H, I, J, K, L, O, P, AD, AV, AG, AW, AH, AX, BB, BD, BF, AT

        for _, row in santom_df.iterrows():
            record = {}
            
            # Copy all columns from santom_df to record
            for col in santom_df.columns:
                record[col] = row[col]
                # try:
                #     float_val = float(row[col])
                #     if float_val.is_integer():
                #         record[col] = int(float_val)  # 123.0 → 123
                #     else:
                #         record[col] = float_val       # 123.45 → 123.45
                # except:
                #     record[col] = row[col]            # Keep as string if conversion fails

            # Check if Account Type is "P" and perform special processing
            if G in santom_df.columns and row[G] == "P":
                # For Account Type "P", calculate balance and assign to AV
                if AD in santom_df.columns:
                    try:
                        balance = float(row[AD]) - float(cash_with_ncl or 0)
                        record[AV] = balance
                    except (ValueError, TypeError):
                        record[AV] = 0
                else:
                    record[AV] = 0
            else:
                # For other account types, copy data from specific columns
                if AD in santom_df.columns:
                    record[AV] = row[AD]
                if AG in santom_df.columns:
                    record[AW] = row[AG]
                if AX in santom_df.columns:
                    record[AH] = row[AX]

            data.append(record)
        return data

    def _create_zip_and_save(self, cash_collateral_cds, cash_collateral_fno, daily_margin_nsecr, 
                           daily_margin_nsefno, x_cp_master, f_cp_master, cd_collateral_valuation_lookup, fo_collateral_valuation_lookup,
                            #  sec_pledge, 
                             output_file, output_path):
        """Create ZIP file and save to database"""
        # Create ZIP file
        dt = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"SEGREGATION_REPORT_{dt}.zip"
        zip_path = os.path.join(output_path, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            # Add input files
            input_files = [
                (cash_collateral_cds, "CashCollateral_CDS.xls"),
                (cash_collateral_fno, "CashCollateral_FNO.xls"),
                (daily_margin_nsecr, "Daily_Margin_Report_NSECR.xls"),
                (daily_margin_nsefno, "Daily_Margin_Report_NSEFNO.xls"),
                (x_cp_master, "X_CPMaster_data.xlsx"),
                (f_cp_master, "F_CPMaster_data.xlsx"),
                (cd_collateral_valuation_lookup, "Collateral Valuation Report_cds.xls"),
                (fo_collateral_valuation_lookup, "Collateral Valuation Report_fno.xls"),
                # (sec_pledge, "F_90123_SEC_PLEDGE_09092025_02.csv.gz")
            ]
            
            for file_path, arcname in input_files:
                zipf.write(file_path, arcname)
            
            # Add output file
            zipf.write(output_file, os.path.basename(output_file))
        
        # Insert into database
        with open(zip_path, 'rb') as f:
            zip_blob = f.read()
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        insert_report(self.db_path, report_type="SEGREGATION_REPORT", 
                     created_at=timestamp, modified_at=timestamp, report_blob=zip_blob)
    
    def _get_master_records(self, av=False, at=False, all_records=False):
        """
        Simple function to get master records based on flags
        
        Args:
            av (bool): Return AV_Records if True
            at (bool): Return AT_Records if True  
            all_records (bool): Return all records combined if True
            
        Returns:
            list: Requested records
        """
        import json
        import os
        
        master_records_json_path = "master_records.json"
        av_records = []
        at_records = []

        if os.path.exists(master_records_json_path):
            try:
                with open(master_records_json_path, 'r') as f:
                    all_master_data = json.load(f)

                av_records = all_master_data.get('AV_Records', [])
                at_records = all_master_data.get('AT_Records', [])

                # print(f"✅ Loaded {len(av_records)} AV records and {len(at_records)} AT records")

            except Exception as e:
                print(f"❌ Error reading master records JSON: {e}")
        
        # Return based on flags
        if av:
            return av_records
        elif at:
            return at_records
        elif all_records:
            return av_records + at_records
        else:
            return av_records, at_records  # Default: return both separately
