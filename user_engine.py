from datetime import date

def get_user_current_data(client, spreadsheet_name, mmyy, user_name):
    try:
        sh = client.open(spreadsheet_name)
        p_ws = sh.worksheet("Partners")
        nl_ws = sh.worksheet("Namelist")
        c_ws = sh.worksheet(f"{mmyy}C")

        # get driving status from Namelist sheet column D
        driving = "NON-DRIVER"
        try:
            nl_cell = nl_ws.find(user_name, in_column=2)
            driving = nl_ws.cell(nl_cell.row, 4).value or "NON-DRIVER"
        except:
            pass

        # get partner from Partners sheet
        # structure: col B = Names (everyone), col C = Partner
        partner = "None"
        try:
            cell = p_ws.find(user_name, in_column=2)
            partner = p_ws.cell(cell.row, 3).value or "None"
        except:
            pass

        # get traits from Namelist sheet column E (same row as driving status)
        traits = ""
        try:
            traits = nl_ws.cell(nl_cell.row, 5).value or ""
            traits = traits.strip()
        except:
            pass

        # get X and D markers
        user_cell = c_ws.find(user_name, in_column=2)
        row_values = c_ws.row_values(user_cell.row)
        c_list = [str(i+1) for i, v in enumerate(row_values[4:35]) if v == 'X']
        p_list = [str(i+1) for i, v in enumerate(row_values[4:35]) if v == 'D']

        return {
            "partner": partner,
            "driving": driving,
            "traits": traits,
            "constraints": ", ".join(c_list),
            "preferences": ", ".join(p_list)
        }
    except:
        return None

def update_user_data(client, spreadsheet_name, mmyy, user_name, partner, driving_status, traits, constraints, preferences, status_string):
    logs = []
    try:
        sh = client.open(spreadsheet_name)
        p_ws = sh.worksheet("Partners")
        nl_ws = sh.worksheet("Namelist")

        # update driving status in Namelist sheet column D
        try:
            nl_cell = nl_ws.find(user_name, in_column=2)
            nl_ws.update_acell(f'D{nl_cell.row}', driving_status)
            logs.append(f"✅ Step 1a: Updated driving status in Namelist for {user_name}")
        except:
            logs.append(f"⚠️ Could not update driving status in Namelist for {user_name}")

        # update traits in Namelist sheet column E
        try:
            nl_cell = nl_ws.find(user_name, in_column=2)
            nl_ws.update_acell(f'E{nl_cell.row}', traits.strip() if traits else "")
            logs.append(f"✅ Step 1a-traits: Updated traits in Namelist for {user_name}")
        except:
            logs.append(f"⚠️ Could not update traits in Namelist for {user_name}")

        # update Partners sheet
        # structure: col B = Names (everyone), col C = Partner
        # read all partner data once to avoid repeated API calls
        p_data = p_ws.get_all_values()
        name_to_prow = {}
        for i, row in enumerate(p_data):
            n = row[1].strip() if len(row) > 1 else ""
            if n:
                name_to_prow[n] = i + 1  # 1-indexed gspread row

        partner_name = partner if partner != "None" else ""
        p_updates = []

        if user_name in name_to_prow:
            u_prow = name_to_prow[user_name]

            # read old partner before overwriting
            old_partner = p_data[u_prow - 1][2].strip() if len(p_data[u_prow - 1]) > 2 else ""

            # write new partner for user
            p_updates.append({'range': f'C{u_prow}', 'values': [[partner_name]]})

            # blank old partner's col C if they had a different partner
            if old_partner and old_partner != partner_name and old_partner in name_to_prow:
                old_prow = name_to_prow[old_partner]
                p_updates.append({'range': f'C{old_prow}', 'values': [[""]]})

            # update new partner's row
            if partner_name and partner_name in name_to_prow:
                np_row = name_to_prow[partner_name]
                np_old_partner = p_data[np_row - 1][2].strip() if len(p_data[np_row - 1]) > 2 else ""
                # blank new partner's old partner if different
                if np_old_partner and np_old_partner != user_name and np_old_partner in name_to_prow:
                    npo_row = name_to_prow[np_old_partner]
                    p_updates.append({'range': f'C{npo_row}', 'values': [[""]]})
                # write user as new partner's partner
                p_updates.append({'range': f'C{np_row}', 'values': [[user_name]]})

            if p_updates:
                p_ws.batch_update(p_updates, value_input_option='USER_ENTERED')
            logs.append(f"✅ Step 1b: Updated Partners sheet for {user_name} → {partner_name or 'None'}")

        else:
            logs.append(f"⚠️ {user_name} not found in Partners sheet — skipping partner update.")

        # update constraint sheet
        c_ws = sh.worksheet(f"{mmyy}C")
        user_cell = c_ws.find(user_name, in_column=2)
        u_row = user_cell.row
        date_start_col = 5   # col E = 5 in 1-indexed gspread
        date_end_col = 35    # col AI = 35 in 1-indexed gspread
        c_updates = []

        c_updates.append({'range': f'AR{u_row}', 'values': [[status_string]]})  # status = AR

        def get_col_let(n):
            if n <= 26:
                return chr(64 + n)
            else:
                first = chr(64 + (n - 1) // 26)
                second = chr(64 + (n - 1) % 26 + 1)
                return f"{first}{second}"

        # wipe existing data
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
    if not day_string: return []
    days = []
    mm = int(month_year_str[:2])
    yy = 2000 + int(month_year_str[2:])

    parts = [p.strip() for p in str(day_string).split(",")]
    for p in parts:
        if p.isdigit():
            days.append(date(yy, mm, int(p)))
    return days

def format_date_list(history_collection):
    day_nums = set()
    for item in history_collection:
        if hasattr(item, 'day'):
            day_nums.add(item.day)
        elif isinstance(item, int):
            day_nums.add(item)

    return ", ".join(map(str, sorted(list(day_nums))))

def calendar_view(client, spreadsheet_name, mmyy):
    try:
        sh = client.open(spreadsheet_name)

        # 1. Get all available sheet titles to check for existence
        all_sheet_titles = [s.title for s in sh.worksheets()]

        # 2. Define our targets
        primary_sheet = f"{mmyy}D"
        backup_sheet = f"{mmyy}C"

        # 3. If-Else Logic for sheet selection
        if primary_sheet in all_sheet_titles:
            worksheet = sh.worksheet(primary_sheet)
            sheet_used = "D"
        elif backup_sheet in all_sheet_titles:
            worksheet = sh.worksheet(backup_sheet)
            sheet_used = "C"
        else:
            return None, None, f"No data for this {mmyy} was found."

        # 4. Pull data once the correct sheet is selected
        raw_data = worksheet.get_all_values()

        # Data starts from index 3 (Row 4)
        rows = raw_data[3:]

        # Initialize dictionary for 31 days
        roster = {str(day): {"duty": [], "standby": []} for day in range(1, 32)}

        for row in rows:
            # Safety check: skip empty rows or rows without names
            if not row or len(row) < 2 or not row[1].strip():
                continue

            name = row[1].strip()

            # Iterate through 31 columns (Starting from Column E / Index 4)
            for i in range(31):
                col_idx = 4 + i
                if col_idx >= len(row):
                    break

                day_num = str(i + 1)
                planned_status = row[col_idx].strip().upper()

                if day_num in roster:
                    if planned_status == 'D':
                        roster[day_num]["duty"].append(name)
                    elif planned_status == 'S':
                        roster[day_num]["standby"].append(name)

        return roster, sheet_used, None

    except Exception as e:
        return None, None, str(e)