import gspread

def get_namelist(client, spreadsheet_name):
    """Fetches the list of names from the 'Namelist' tab."""
    try:
        sh = client.open(spreadsheet_name)
        ws = sh.worksheet("Namelist")
        records = ws.get_all_records()
        return [r['NAME'] for r in records if r.get('NAME')]
    except Exception as e:
        print(f"Error fetching namelist: {e}")
        return []

def get_person_driving_status(p_ws, name):
    """Helper to find any person's status in the Partners sheet."""
    if not name or name == "None": return ""
    try:
        # check first name (column B, status in column C)
        cell = p_ws.find(name, in_column=2)
        return p_ws.cell(cell.row, 3).value
    except:
        try:
            # check first name (column D, status in column E)
            cell = p_ws.find(name, in_column=4)
            return p_ws.cell(cell.row, 5).value
        except:
            return "NON-DRIVER"

def get_user_current_data(client, spreadsheet_name, mmyy, user_name):
    """FETCH PREVIOUS DATA: Scans sheets to pre-fill the form."""
    try:
        sh = client.open(spreadsheet_name)
        p_ws = sh.worksheet("Partners")
        c_ws = sh.worksheet(f"{mmyy}C")
        
        # get partner and driving information
        partner, driving = "None", "NON-DRIVER"
        try:
            cell = p_ws.find(user_name, in_column=2) # search column B
            partner = p_ws.cell(cell.row, 4).value or "None"
            driving = p_ws.cell(cell.row, 3).value or "NON-DRIVER"
        except:
            try:
                cell = p_ws.find(user_name, in_column=4) # search column D
                partner = p_ws.cell(cell.row, 2).value or "None"
                driving = p_ws.cell(cell.row, 5).value or "NON-DRIVER"
            except: pass

        # get X and D markers
        user_cell = c_ws.find(user_name, in_column=2)
        row_values = c_ws.row_values(user_cell.row)
        # column F is index 5, days 1-31
        c_list = [str(i+1) for i, v in enumerate(row_values[5:36]) if v == 'X']
        p_list = [str(i+1) for i, v in enumerate(row_values[5:36]) if v == 'D']

        return {
            "partner": partner,
            "driving": driving,
            "constraints": ", ".join(c_list),
            "preferences": ", ".join(p_list)
        }
    except:
        return None

def update_user_data(client, spreadsheet_name, mmyy, user_name, partner, driving_status, constraints, preferences, status_string):
    logs = []
    try:
        sh = client.open(spreadsheet_name)
        p_ws = sh.worksheet("Partners")
        
        # find user in partners sheet
        cell = None
        is_primary = True
        try:
            cell = p_ws.find(user_name, in_column=2)
            is_primary = True
        except gspread.exceptions.CellNotFound:
            try:
                cell = p_ws.find(user_name, in_column=4)
                is_primary = False
            except gspread.exceptions.CellNotFound:
                logs.append(f"⚠️ {user_name} not found in Partners sheet.")

        if cell:
            row_idx = cell.row
            partner_name = partner if partner != "None" else ""
            partner_driving_status = get_person_driving_status(p_ws, partner_name)
            
            if is_primary:
                # user in B: update C (user status) and D/E (partner mame/partner status)
                p_updates = [
                    {'range': f'C{row_idx}', 'values': [[driving_status]]},
                    {'range': f'D{row_idx}', 'values': [[partner_name]]},
                    {'range': f'E{row_idx}', 'values': [[partner_driving_status]]}
                ]
            else:
                # user in D: update E (user status) and B/C (partner name/partner status)
                p_updates = [
                    {'range': f'E{row_idx}', 'values': [[driving_status]]},
                    {'range': f'B{row_idx}', 'values': [[partner_name]]},
                    {'range': f'C{row_idx}', 'values': [[partner_driving_status]]}
                ]
            
            p_ws.batch_update(p_updates, value_input_option='USER_ENTERED')
            logs.append(f"✅ Step 1: Updated Partners sheet for {user_name} and partner {partner_name}")

        # update sheet
        c_ws = sh.worksheet(f"{mmyy}C")
        user_cell = c_ws.find(user_name, in_column=2)
        u_row = user_cell.row
        date_start_col = 6 
        date_end_col = 36
        c_updates = []

        c_updates.append({'range': f'AS{u_row}', 'values': [[status_string]]})

        def get_col_let(n):
            if n <= 26:
                return chr(64 + n)
            else:
                first = chr(64 + (n - 1) // 26)
                second = chr(64 + (n - 1) % 26 + 1)
                return f"{first}{second}"
            
        # wipes existing data in the sheet
        clear_range = f"{get_col_let(date_start_col)}{u_row}:{get_col_let(date_end_col)}{u_row}"
        blank_row = [["" for _ in range(31)]]
        c_ws.update(clear_range, blank_row)

        if constraints:
            for day in [d.strip() for d in constraints.split(',') if d.strip().isdigit()]:
                col_let = get_col_let(date_start_col + int(day) - 1)
                c_updates.append({'range': f'{col_let}{u_row}', 'values': [['X']]})
        
        if preferences:
            for day in [d.strip() for d in preferences.split(',') if d.strip().isdigit()]:
                col_let = get_col_let(date_start_col + int(day) - 1)
                c_updates.append({'range': f'{col_let}{u_row}', 'values': [['D']]})

        if c_updates:
            c_ws.batch_update(c_updates, value_input_option='USER_ENTERED')
            logs.append(f"✅ Step 2: Updated {mmyy}C markers.")
        
        return True, logs

    except Exception as e:
        return False, [f"❌ Error: {str(e)}"]
    
def parse_string_to_days(day_string, month_year_str):
        """Converts '1, 2' string into a list of date objects for the calendar memory."""
        if not day_string: return []
        days = []
        # month_year_str is 'MMYY' -> e.g., '0126'
        mm = int(month_year_str[:2])
        yy = 2000 + int(month_year_str[2:])
        
        parts = [p.strip() for p in str(day_string).split(",")]
        for p in parts:
            if p.isdigit():
                days.append(date(yy, mm, int(p)))
        return days

def format_date_list(history_collection):
        """Safely converts a collection of dates OR ints into a sorted string."""
        day_nums = set()
        for item in history_collection:
            if hasattr(item, 'day'):
                day_nums.add(item.day)
            elif isinstance(item, int):
                day_nums.add(item)
        
        return ", ".join(map(str, sorted(list(day_nums))))