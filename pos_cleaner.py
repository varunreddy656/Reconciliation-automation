import pandas as pd
import io
import re
import os
import shutil
import tempfile
from datetime import datetime
from werkzeug.utils import secure_filename

def clean_pos_excel(file_storage):
    """
    Detects POS type (Posist/Petpooja) and returns a standardized cleaned dataframe.
    """
    try:
        file_bytes = file_storage.read()
        file_storage.seek(0)
        
        # STEP 1 — DETECT POS TYPE
        df_detect = pd.read_excel(io.BytesIO(file_bytes), nrows=10, header=None)
        
        pos_type = "unknown"
        for i in range(min(10, len(df_detect))):
            row_str = " ".join(df_detect.iloc[i].astype(str).tolist()).lower()
            if any(kw in row_str for kw in ["from :", "transaction date", "order no"]):
                pos_type = "posist"
                break
            if any(kw in row_str for kw in ["order report", "payment wise", "invoice"]):
                pos_type = "petpooja"
                break
            
        print(f"--- Detected POS Type: {pos_type.upper()} ---")
        
        cleaned_df = pd.DataFrame()
        
        if pos_type == "petpooja":
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
            header_idx = -1
            # Search for the header row containing 'Invoice' and 'Date' and 'Payment Type'
            for idx, row in df_raw.iterrows():
                row_vals = " ".join(row.astype(str).tolist()).lower()
                if "invoice" in row_vals and "date" in row_vals and "payment type" in row_vals:
                    header_idx = idx
                    break
            
            if header_idx != -1:
                # Read again using header_idx as header
                df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_idx)
                
                # Cleanup noise immediately below headers: Total, Min, Max, Avg, Count rows
                # Check the first column (Invoice No.)
                noise_pattern = "Total|Min\.|Max\.|Avg\.|Count|Grand Total"
                df = df[~df.iloc[:, 0].astype(str).str.contains(noise_pattern, na=False, case=False)]
                
                # Ensure data starts correctly (remove any remaining noise)
                # Actual data rows usually have a date or invoice number
                df = df[df.iloc[:, 1].notnull()] # Date column should not be null

                # Column standardization mapping
                col_map = {}
                for col in df.columns:
                    c_low = str(col).lower().strip()
                    if "invoice" in c_low: col_map[col] = "Invoice_No"
                    elif "date" == c_low: col_map[col] = "Transaction_Date"
                    elif "payment" in c_low and "type" in c_low: col_map[col] = "Payment_Type"
                    elif "order" in c_low and "type" in c_low: col_map[col] = "Order_Type"
                    elif "status" == c_low: col_map[col] = "Status"
                    elif "area" in c_low: col_map[col] = "Area"
                
                df = df.rename(columns=col_map)
                
                # Ensure crucial columns exist
                if 'Transaction_Date' in df.columns:
                    df['Transaction_Date'] = pd.to_datetime(df['Transaction_Date'], errors='coerce')
                    df = df[df['Transaction_Date'].notnull()]
                
                # Standardize Amount Columns
                amt_cols = ['Cash', 'Card', 'Due Payment', 'Other', 'Wallet', 'Online']
                for col in amt_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                cleaned_df = df
            else:
                pos_type = "unknown"

        elif pos_type == "posist":
            df = pd.read_excel(io.BytesIO(file_bytes), skiprows=1)
            cols_mapping = {
                'Transaction Date': 'Transaction_Date',
                'Payment T': 'Payment_Type',
                'Payment M': 'Payment_Method',
                'Amount': 'Amount'
            }
            df = df.rename(columns=cols_mapping)
            if 'Transaction_Date' in df.columns:
                df['Transaction_Date'] = pd.to_datetime(df['Transaction_Date'], errors='coerce')
            if 'Amount' in df.columns:
                df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
            cleaned_df = df

        if pos_type == "unknown":
            cleaned_df = pd.read_excel(io.BytesIO(file_bytes))
            # Try basic date detection
            for col in cleaned_df.columns:
                if "date" in str(col).lower():
                    cleaned_df[col] = pd.to_datetime(cleaned_df[col], errors='coerce')
                    if "Transaction_Date" not in cleaned_df.columns:
                        cleaned_df.rename(columns={col: "Transaction_Date"}, inplace=True)

        return {
            "pos_type": pos_type,
            "dataframe": cleaned_df
        }

    except Exception as e:
        print(f"❌ Error in clean_pos_excel: {str(e)}")
        return {"pos_type": "unknown", "dataframe": pd.DataFrame()}

def aggregate_pos(df, pos_type, week_ranges=None, custom_dinein_ranges=None):
    """
    Performs hardcoded aggregation based on POS type and user-provided week ranges.
    """
    try:
        if df.empty:
            return {"weeks": [], "channels": [], "message": "No data found"}

        # Initialize default weeks
        weeks = ["Week 1 (1st-7th)", "Week 2 (8th-14th)", "Week 3 (15th-21st)", "Week 4 (22nd-28th)", "Week 5 (29th+)"]
        
        f_start = f_end = l_start = l_end = None
        if week_ranges:
            try:
                # Convert string inputs to date objects
                f_start = pd.to_datetime(week_ranges.get('f_start')).date() if week_ranges.get('f_start') else None
                f_end = pd.to_datetime(week_ranges.get('f_end')).date() if week_ranges.get('f_end') else None
                l_start = pd.to_datetime(week_ranges.get('l_start')).date() if week_ranges.get('l_start') else None
                l_end = pd.to_datetime(week_ranges.get('l_end')).date() if week_ranges.get('l_end') else None
                
                # Update labels dynamically using Rolling 7-day logic
                if f_start and f_end:
                    w1_e = f_end.day
                    w5_s = l_start.day if l_start else 29
                    
                    weeks[0] = f"Week 1 ({f_start.day}{'st' if f_start.day==1 else 'th'}-{w1_e}th)"
                    weeks[1] = f"Week 2 ({w1_e + 1}th-{w1_e + 7}th)"
                    weeks[2] = f"Week 3 ({w1_e + 8}th-{w1_e + 14}th)"
                    weeks[3] = f"Week 4 ({w1_e + 15}th-{w5_s - 1}th)"
                
                if l_start and l_end:
                    weeks[4] = f"Week 5 ({l_start.day}th-{l_end.day}th)"
                    
            except Exception as week_err:
                print(f"⚠️ Week Range Parsing Error: {week_err}")

        def get_label(dt, s_start, s_end, e_start, e_end):
            if pd.isna(dt): return "Unknown"
            d = dt.date() if hasattr(dt, 'date') else dt
            day = d.day
            
            # Dynamic Boundaries
            w1_end = s_end.day if s_end else 7
            w5_start = e_start.day if e_start else 29
            
            # Rolling 7-day pattern for middle weeks
            w2_end = w1_end + 7
            w3_end = w2_end + 7
            
            if day <= w1_end: return weeks[0]
            if day <= w2_end: return weeks[1]
            if day <= w3_end: return weeks[2]
            if day < w5_start: return weeks[3]
            return weeks[4]

        # Pre-calculate labels using separate logic for Global and Swiggy
        df['Week_Global'] = df['Transaction_Date'].apply(lambda x: get_label(x, f_start, f_end, l_start, l_end))
        
        sw_f_s = sw_f_e = sw_l_s = sw_l_e = None
        if week_ranges and week_ranges.get('swiggy'):
            sw = week_ranges['swiggy']
            try:
                # Use flexible parsing for Swiggy specific ranges
                if sw.get('f_start'): sw_f_s = pd.to_datetime(sw['f_start']).date()
                if sw.get('f_end'): sw_f_e = pd.to_datetime(sw['f_end']).date()
                if sw.get('l_start'): sw_l_s = pd.to_datetime(sw['l_start']).date()
                if sw.get('l_end'): sw_l_e = pd.to_datetime(sw['l_end']).date()
                print(f"DEBUG: Swiggy Ranges Parsed - F_End: {sw_f_e}, L_Start: {sw_l_s}")
            except Exception as e:
                print(f"⚠️ Swiggy Week Range Parsing Error: {e}")
        
        df['Week_Swiggy'] = df['Transaction_Date'].apply(lambda x: get_label(x, sw_f_s, sw_f_e, sw_l_s, sw_l_e))

        results = {"weeks": weeks, "channels": []}

        if pos_type == "petpooja":
            def detect_channel(row):
                area = str(row.get('Area', '')).lower()
                p_type = str(row.get('Payment_Type', '')).lower()
                o_type = str(row.get('Order_Type', '')).lower()
                
                # Zomato Dine-In
                if 'dine' in o_type and ('zomato' in p_type or 'zpay' in p_type): return 'Zomato Dine-In'
                
                # Swiggy Dineout / SD
                if 'sd' in p_type or 'dine out' in p_type or 'swiggy dineout' in p_type: return 'Swiggy Dineout'
                if 'eazydiner' in p_type: return 'EazyDiner'
                
                if 'zomato' in area or 'zomato' in p_type or 'zomato' in o_type: return 'Zomato'
                if 'swiggy' in area or 'swiggy' in p_type or 'swiggy' in o_type: return 'Swiggy'
                if 'dine out' in p_type or 'dineout' in p_type or 'dine out' in o_type: return 'Dineout'
                if 'magicpin' in p_type: return 'MagicPin'
                if any(kw in o_type for kw in ['dine', 'pick', 'take', 'parcel']): return 'Dine In'
                return o_type.title() or 'Other'

            df['Detected_Channel'] = df.apply(detect_channel, axis=1)
            
            for channel in df['Detected_Channel'].unique():
                if not channel or str(channel).lower() in ['nan']: continue
                
                is_delivery = channel in ['Zomato', 'Swiggy']
                is_zpay = channel == 'Zomato Dine-In'
                is_custom_dinein = channel in ['Swiggy Dineout', 'EazyDiner']
                
                channel_obj = {"name": channel, "payment_methods": []}
                week_col = 'Week_Swiggy' if channel == 'Swiggy' else 'Week_Global'

                if is_custom_dinein and custom_dinein_ranges:
                    # Specialized logic for Swiggy Dineout / EazyDiner using AI ranges
                    channel_obj["is_custom"] = True
                    channel_obj["periods"] = [r['label'] for r in custom_dinein_ranges]
                    
                    p_data = {}
                    df_success = df[(df['Detected_Channel'] == channel) & 
                                    (df['Status'].astype(str).str.lower().str.contains('success', na=False))]
                    
                    for r in custom_dinein_ranges:
                        # Filter by start/end day in the month
                        mask = (df_success['Transaction_Date'].dt.day >= r['start_day']) & \
                               (df_success['Transaction_Date'].dt.day <= r['end_day'])
                        val = df_success[mask]['Other'].sum() if 'Other' in df_success.columns else 0
                        p_data[r['label']] = {"value": float(val)}
                    
                    if any(v['value'] > 0 for v in p_data.values()):
                        channel_obj["payment_methods"].append({
                            "method": channel.upper(),
                            "status_label": "SUCCESS",
                            "data": p_data
                        })
                
                elif is_delivery:
                    online_data = {}
                    for week in weeks:
                        mask = (df[week_col] == week) & (df['Detected_Channel'] == channel)
                        p_mask = mask & (df['Payment_Type'].astype(str).str.lower().str.contains('online', na=False))
                        val = df[p_mask]['Online'].sum() if 'Online' in df.columns else 0
                        online_data[week] = {"value": float(val)}
                    
                    if any(v['value'] > 0 for v in online_data.values()):
                        channel_obj["payment_methods"].append({
                            "method": "ONLINE", 
                            "status_label": "SUCCESS & CANCELLED",
                            "data": online_data
                        })
                
                elif is_zpay:
                    df_success = df[df['Status'].astype(str).str.lower().str.contains('success', na=False)]
                    zpay_data = {}
                    for week in weeks:
                        mask = (df_success[week_col] == week) & (df_success['Detected_Channel'] == channel)
                        val = df_success[mask]['Other'].sum() if 'Other' in df_success.columns else 0
                        zpay_data[week] = {"value": float(val)}
                    if any(v['value'] > 0 for v in zpay_data.values()):
                        channel_obj["payment_methods"].append({
                            "method": "ZOMATO PAY",
                            "status_label": "SUCCESS",
                            "data": zpay_data
                        })
                
                else:
                    df_success = df[df['Status'].astype(str).str.lower().str.contains('success', na=False)]
                    
                    # UPI
                    upi_data = {}
                    has_upi = False
                    for week in weeks:
                        mask = (df_success[week_col] == week) & (df_success['Detected_Channel'] == channel)
                        upi_mask = mask & (df_success['Payment_Type'].astype(str).str.lower().str.contains('upi|part payment', na=False))
                        val = df_success[upi_mask]['Other'].sum() if 'Other' in df_success.columns else 0
                        upi_data[week] = {"value": float(val)}
                        if val > 0: has_upi = True
                    if has_upi: channel_obj["payment_methods"].append({"method": "UPI", "data": upi_data})
                    
                    # CARD
                    card_data = {}
                    has_card = False
                    for week in weeks:
                        mask = (df_success[week_col] == week) & (df_success['Detected_Channel'] == channel)
                        val = df_success[mask]['Card'].sum() if 'Card' in df_success.columns else 0
                        card_data[week] = {"value": float(val)}
                        if val > 0: has_card = True
                    if has_card: channel_obj["payment_methods"].append({"method": "CARD", "data": card_data})

                if channel_obj["payment_methods"]:
                    results["channels"].append(channel_obj)

        elif pos_type == "posist":
            # POSist basic logic
            if 'Payment_Method' in df.columns:
                for method in df['Payment_Method'].unique():
                    if not method or str(method).lower() == 'nan': continue
                    ch_obj = {"name": str(method), "show_cancelled": False, "payment_methods": [{"method": "Total", "data": {}}]}
                    for week in weeks:
                        val = df[(df['Week_Label'] == week) & (df['Payment_Method'] == method)]['Amount'].sum()
                        ch_obj["payment_methods"][0]["data"][week] = {"success": float(val), "cancelled": 0}
                    if any(v['success'] > 0 for v in ch_obj["payment_methods"][0]["data"].values()):
                        results["channels"].append(ch_obj)

        return results
    except Exception as e:
        print(f"❌ Error in aggregate_pos: {str(e)}")
        import traceback
        traceback.print_exc()
        return {"weeks": [], "channels": [], "message": str(e)}

def generate_pos_report(data, output_path):
    """Generates an Excel report from aggregated POS data"""
    try:
        rows = []
        weeks = data['weeks']
        for channel in data['channels']:
            for method in channel['payment_methods']:
                row_success = {"Channel": channel['name'], "Payment Method": method['method'], "Status": "SUCCESS"}
                for week in weeks:
                    w_data = method['data'].get(week, {"success": 0, "cancelled": 0})
                    row_success[week] = w_data['success']
                rows.append(row_success)
                    
        if rows:
            df_out = pd.DataFrame(rows)
            cols = ["Channel", "Payment Method", "Status"] + [w for w in weeks if w in df_out.columns]
            df_out[cols].to_excel(output_path, index=False)
            return True
        else:
            # Fallback for empty
            df_out = pd.DataFrame(columns=["Channel", "Payment Method", "Status"] + weeks)
            df_out.to_excel(output_path, index=False)
            return True
    except Exception as e:
        print(f"❌ Error generating POS report: {str(e)}")
        return False
