
import os
import openpyxl
import re
from datetime import datetime
from werkzeug.utils import secure_filename
import gc

def get_week_ranges(first_start, first_end, last_start, last_end):
    """Calculates 5 week ranges based on the first and last week input."""
    try:
        f_s, f_e = int(float(first_start)), int(float(first_end))
        l_s, l_e = int(float(last_start)), int(float(last_end))
        
        # Week 1: f_s to f_e
        # Week 2: f_e + 1 to f_e + 7
        # Week 3: f_e + 8 to f_e + 14
        # Week 4: f_e + 15 to f_e + 21
        # Week 5: l_s to l_e
        
        return [
            (f_s, f_e),
            (f_e + 1, f_e + 7),
            (f_e + 8, f_e + 14),
            (f_e + 15, f_e + 21),
            (l_s, l_e)
        ], [f_s, f_e, f_e+1, f_e+7, f_e+8, f_e+14, f_e+15, f_e+21, l_s, l_e]
    except:
        return [], []

def safe_float(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    try:
        clean_v = re.sub(r'[^\d.]', '', str(v))
        return float(clean_v) if clean_v else 0.0
    except:
        return 0.0

def process_zomato_pay(invoice_files, template_path, output_dir, update_progress=None, 
                       client_name="", month="", first_start=None, first_end=None, 
                       last_start=None, last_end=None):
    """
    Ultra-optimized Zomato Pay reconciliation with refined logic.
    """
    temp_files = []
    try:
        if update_progress: update_progress(5)

        # 1. Prepare Output Workbook
        if not os.path.exists(template_path):
            return None, f"Template file not found at {template_path}"
        
        out_wb = openpyxl.load_workbook(template_path)
        ws_calc = out_wb["Zpay Calculations"] if "Zpay Calculations" in out_wb.sheetnames else out_wb.create_sheet("Zpay Calculations")
        ws_ads = out_wb["Zpay Ads"] if "Zpay Ads" in out_wb.sheetnames else out_wb.create_sheet("Zpay Ads")
        
        # Clear existing content
        ws_calc.delete_rows(1, ws_calc.max_row)
        ws_ads.delete_rows(1, ws_ads.max_row)

        processed_data = [] # Stores Transaction Summary
        ads_data = [] # Stores Ad Summary

        # 2. Fast Input Reading (Read-Only)
        for idx, file in enumerate(invoice_files):
            filename = secure_filename(file.filename)
            temp_path = os.path.join(output_dir, f"temp_zpay_{filename}")
            file.save(temp_path)
            temp_files.append(temp_path)

            wb_in = openpyxl.load_workbook(temp_path, read_only=True, data_only=True)
            if "Transactions summary" in wb_in.sheetnames:
                # Capture from actual data row 7
                for row in wb_in["Transactions summary"].iter_rows(min_row=7, values_only=True):
                    if any(row): processed_data.append(row)
            
            if "Additions & deductions" in wb_in.sheetnames:
                for row in wb_in["Additions & deductions"].iter_rows(min_row=3, values_only=True):
                    if any(row): ads_data.append(row)
            wb_in.close()
        
        if update_progress: update_progress(35)

        # 3. Batch Write to Template (Fast Append)
        # Create 14 row gap as requested
        for _ in range(14): ws_calc.append([])
        for row in processed_data: ws_calc.append(row)
        
        for _ in range(5): ws_ads.append([])
        for row in ads_data: ws_ads.append(row)

        if update_progress: update_progress(60)

        # 4. Calculation Mapping (Headers sit at Row 15)
        headers = [str(cell.value).strip().lower() if cell.value else "" for cell in ws_calc[15]]
        def find_col(possible_names):
            for name in possible_names:
                for idx, h in enumerate(headers):
                    if name.lower() in h: return idx
            return -1

        col_date = find_col(["date and time"])
        col_bill = find_col(["bill amount"])
        col_discount = find_col(["instant discount"])
        col_promo = find_col(["promo share"])
        col_comm = find_col(["commission amount"])
        col_tip = find_col(["tips"])
        col_net = find_col(["net receivable"])

        if col_date == -1 or col_bill == -1:
            return None, "Required date or bill columns missing in Transactions summary."

        # Strict Month Filtering
        month_map = {"january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
                     "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12}
        target_month_num = month_map.get(month.lower())

        weeks, _ = get_week_ranges(first_start, first_end, last_start, last_end)
        weekly_stats = {i: {'bill':0, 'disc':0, 'comm':0, 'tip':0, 'net':0} for i in range(len(weeks))}
        
        # Adjustments for prev/next month
        adj_prev_month = 0.0
        adj_next_month = 0.0

        # Performance: Loop over processed_data list directly (starting from Row 16 equivalent)
        for idx, row in enumerate(processed_data):
            date_val = row[col_date]
            if not date_val: continue
            
            day, m_num = None, None
            if isinstance(date_val, datetime):
                day, m_num = date_val.day, date_val.month
            else:
                parts = re.findall(r'\d+', str(date_val))
                if len(parts) >= 3:
                    if len(parts[0]) == 4: # YYYY-MM-DD
                        day, m_num = int(parts[2]), int(parts[1])
                    else: # DD-MM-YYYY
                        day, m_num = int(parts[0]), int(parts[1])

            if day is None: continue
            
            # Handle Adjustments (Prev/Next month)
            if target_month_num:
                # Logic to determine if prev or next month relative to target
                is_prev = False
                is_next = False
                if m_num < target_month_num:
                    if not (target_month_num == 1 and m_num == 12): is_prev = True # e.g., target Jan (1), m_num Dec (12) -> prev
                    else: is_next = True # e.g., target Feb (2), m_num Jan (1) -> prev
                elif m_num > target_month_num:
                    if not (target_month_num == 12 and m_num == 1): is_next = True # e.g., target Dec (12), m_num Jan (1) -> next
                    else: is_prev = True # e.g., target Nov (11), m_num Dec (12) -> next
                
                if is_prev: adj_prev_month += safe_float(row[col_net])
                if is_next: adj_next_month += safe_float(row[col_net])

                # Skip weekly distribution for adjustment rows
                if m_num != target_month_num: continue

            for i, (ws, we) in enumerate(weeks):
                if ws <= day <= we:
                    stats = weekly_stats[i]
                    stats['bill'] += safe_float(row[col_bill])
                    # Fixed Logic: Discounts use direct sum to match yellow cell
                    stats['disc'] += (safe_float(row[col_discount]) if col_discount != -1 else 0) + \
                                     (safe_float(row[col_promo]) if col_promo != -1 else 0)
                    stats['comm'] += safe_float(row[col_comm])
                    stats['tip'] += safe_float(row[col_tip]) if col_tip != -1 else 0
                    stats['net'] += safe_float(row[col_net])
                    
                    # Mark week for debugger
                    ws_calc.cell(row=15+idx, column=len(row)+1).value = f"W{i+1}"
                    break

        # 5. Inject Weekly Results into Row 2-6 (G onwards)
        calc_results = {i: stats for i, stats in weekly_stats.items()}
        for i in range(len(weeks)):
            stats = weekly_stats[i]
            x_col = 7 + i # G=7, H=8, etc.
            ws_calc.cell(row=1, column=x_col).value = f"Week {i+1} Recon"
            ws_calc.cell(row=2, column=x_col).value = stats['bill'] * (100.0/105.0)
            ws_calc.cell(row=3, column=x_col).value = stats['disc'] * (100.0/105.0) # Apply 100/105 to Discounts
            ws_calc.cell(row=4, column=x_col).value = stats['comm'] * 1.18
            ws_calc.cell(row=5, column=x_col).value = stats['tip']
            ws_calc.cell(row=6, column=x_col).value = stats['net']

        # 6. Zpay Ads Logic - Correction: headers are in row 6 (after insert_rows(1,5))
        ads_headers = [str(ws_ads.cell(row=6, column=c).value).strip().lower() if ws_ads.cell(row=6, column=c).value else "" for c in range(1, ws_ads.max_column + 1)]
        col_ads_date = -1
        col_ads_amt = -1
        for idx, h in enumerate(ads_headers):
            if "date" in h: col_ads_date = idx
            if "amount" in h: col_ads_amt = idx
        
        ads_weekly = {i: 0.0 for i in range(len(weeks))}
        if col_ads_date != -1 and col_ads_amt != -1:
            for idx, row in enumerate(ads_data):
                date_val = row[col_ads_date]
                if not date_val: continue
                day, m_num = None, None
                if isinstance(date_val, datetime):
                    day, m_num = date_val.day, date_val.month
                else:
                    parts = re.findall(r'\d+', str(date_val))
                    if len(parts) >= 3:
                        if len(parts[0]) == 4: day, m_num = int(parts[2]), int(parts[1])
                        else: day, m_num = int(parts[0]), int(parts[1])
                
                if day is None or (target_month_num and m_num != target_month_num):
                    continue
                
                for i, (ws, we) in enumerate(weeks):
                    if ws <= day <= we:
                        val = safe_float(row[col_ads_amt])
                        ads_weekly[i] += val
                        ws_ads.cell(row=7+idx, column=len(row)+1).value = f"W{i+1}"
                        break
        
        for i in range(len(weeks)):
            ws_ads.cell(row=1, column=7+i).value = f"W{i+1}"
            ws_ads.cell(row=2, column=7+i).value = ads_weekly[i]

        # 7. Final Mapping to Zomato Pay (Consolidated)
        if "Zomato Pay" in out_wb.sheetnames:
            ws_final = out_wb["Zomato Pay"]
            # Set Adjustments first
            ws_final["D23"].value = adj_prev_month
            ws_final["H24"].value = adj_next_month

            for r in range(1, ws_final.max_row + 1):
                label = str(ws_final.cell(row=r, column=3).value or "").strip().lower()
                
                # Mapping Logic
                if "sales (exclusive of gst) before failed and reversed transactions" in label:
                    for i in range(len(weeks)):
                        ws_final.cell(row=r, column=4+i).value = calc_results[i]['bill'] * (100.0/105.0)
                
                elif "less: discounts" in label:
                    for i in range(len(weeks)):
                        ws_final.cell(row=r, column=4+i).value = -(calc_results[i]['disc'] * (100.0/105.0))
                
                elif "add : tips" in label:
                    for i in range(len(weeks)):
                        ws_final.cell(row=r, column=4+i).value = calc_results[i]['tip']
                
                elif "commission (inclusive of gst)" in label:
                    for i in range(len(weeks)):
                        ws_final.cell(row=r, column=4+i).value = calc_results[i]['comm'] * 1.18
                
                elif "zomatopay ads" in label:
                    for i in range(len(weeks)):
                        # Input ads are negative, map as positive: -(-val) = val
                        ws_final.cell(row=r, column=4+i).value = -ads_weekly[i]

        # Save and Cleanup
        output_filename = f"Zomato_Pay_Recon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        full_path = os.path.join(output_dir, output_filename)
        out_wb.save(full_path)
        out_wb.close()
        
        for f in temp_files:
            try: os.remove(f)
            except: pass
        
        if update_progress: update_progress(100)
        gc.collect()
        return output_filename, None

    except Exception as e:
        import traceback; traceback.print_exc()
        return None, str(e)
