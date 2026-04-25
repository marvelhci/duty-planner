from datetime import date

def _col_letter(n):
    """Convert 1-based column index to spreadsheet letter(s), e.g. 1→A, 27→AA."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

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

        # get trait values — any column beyond D (index 4+) is a trait column
        # read headers from row 1 to find trait column positions
        traits = {}
        try:
            all_headers = nl_ws.row_values(1)  # e.g. ['', 'NAME', 'BRANCH', 'DRIVING', 'Seniority', 'Team']
            for col_idx, header in enumerate(all_headers[4:], start=5):  # col 5 onward (1-indexed)
                header = header.strip()
                if header:
                    val = nl_ws.cell(nl_cell.row, col_idx).value or ""
                    traits[header] = val.strip()
        except:
            pass

        # get X and D markers
        user_cell = c_ws.find(user_name, in_column=2)
        row_values = c_ws.row_values(user_cell.row)
        c_list = [str(i+1) for i, v in enumerate(row_values[4:35]) if v == 'X']
        p_list = [str(i+1) for i, v in enumerate(row_values[4:35]) if v == 'D']

        result = {
            "partner": partner,
            "driving": driving,
            "constraints": ", ".join(c_list),
            "preferences": ", ".join(p_list)
        }
        # add trait values as 'trait_<CategoryName>' keys so the form can pre-populate
        for cat, val in traits.items():
            result[f"trait_{cat}"] = val
        return result
    except:
        return None

def update_user_data(client, spreadsheet_name, mmyy, user_name, partner, driving_status, selected_traits, constraints, preferences, status_string):
    """
    selected_traits: dict of { category_name: chosen_option }
                     e.g. {'Seniority': 'Senior', 'Team': 'Alpha'}
                     Written to the matching header columns in the Namelist sheet.
    """
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

        # update trait values — find each category's column by header name in row 1
        if selected_traits:
            try:
                nl_cell = nl_ws.find(user_name, in_column=2)
                all_headers = nl_ws.row_values(1)
                trait_updates = []
                for cat, val in selected_traits.items():
                    # find the column index (1-based) for this category header
                    for col_idx, header in enumerate(all_headers, start=1):
                        if header.strip() == cat:
                            trait_updates.append({
                                'range': f'{_col_letter(col_idx)}{nl_cell.row}',
                                'values': [[val or ""]]
                            })
                            break
                if trait_updates:
                    nl_ws.batch_update(trait_updates, value_input_option='USER_ENTERED')
                logs.append(f"✅ Step 1b: Updated trait values in Namelist for {user_name}")
            except Exception as e:
                logs.append(f"⚠️ Could not update traits in Namelist for {user_name}: {e}")

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

def get_roster_context(client, spreadsheet_name, mmyy):
    """
    Reads {mmyy}C sheet and returns:
      day_duty_counts: { day_int: int }        — total D count per day
      all_person_days: { name_str: [day_int] } — each person's D days
    Day cols are 0-indexed 4..34 (days 1..31). Data rows start at index 3.
    """
    day_duty_counts = {d: 0 for d in range(1, 32)}
    all_person_days = {}
    try:
        sh = client.open(spreadsheet_name)
        all_titles = [s.title for s in sh.worksheets()]
        sheet_name = f"{mmyy}C"
        if sheet_name not in all_titles:
            return day_duty_counts, all_person_days
        ws = sh.worksheet(sheet_name)
        raw = ws.get_all_values()
        for row in raw[3:]:
            if not row or len(row) < 2 or not row[1].strip():
                continue
            name = row[1].strip()
            person_d_days = []
            for i in range(31):
                col_idx = 4 + i
                if col_idx >= len(row):
                    break
                if row[col_idx].strip().upper() == 'D':
                    day_duty_counts[i + 1] += 1
                    person_d_days.append(i + 1)
            all_person_days[name] = person_d_days
    except Exception:
        pass
    return day_duty_counts, all_person_days


def get_applicable_constraints(config):
    """
    Reads CONFIG (as returned by fetch_config) and extracts hard constraints
    enforceable at submission time. Returns a list of dicts with key 'kind':
      kind='day_max':          max D per day            → {limit}
      kind='person_month_max': max D per person/month   → {limit}
      kind='person_week_max':  max D per person/week    → {limit}
      kind='gap_dd':           min gap between own D's  → {days}
    Automatically picks up new constraints added via the Dev page.
    """
    results = []
    for cid, cv in config.items():
        if cid.startswith("_"):
            continue
        if not cv.get("active", True):
            continue
        if cv.get("type", "").lower() != "hard":
            continue
        rule = cv.get("rule", {})
        if not rule:
            continue

        cls   = rule.get("class", "")
        soft  = rule.get("soft", False)
        label = cv.get("label", cid)

        if cls == "value" and not soft:
            subj1 = rule.get("subject1", "person")
            op    = rule.get("operator", "=")
            subj2 = rule.get("subject2", "D").upper()
            per   = rule.get("per", "month")
            if subj2 != "D" or op not in ("=", "<="):
                continue
            try:
                number = int(rule.get("number", 1))
            except Exception:
                continue
            if subj1 == "day":
                results.append({"kind": "day_max", "limit": number, "label": label})
            elif subj1 == "person" and per == "month":
                results.append({"kind": "person_month_max", "limit": number, "label": label})
            elif subj1 == "person" and per == "week":
                results.append({"kind": "person_week_max", "limit": number, "label": label})

        elif cls == "gap" and not soft:
            from_type = rule.get("from_type", "D").upper()
            to_type   = rule.get("to_type", "D").upper()
            if from_type != "D" or to_type != "D":
                continue
            try:
                days = int(rule.get("days", 4))
            except Exception:
                continue
            results.append({"kind": "gap_dd", "days": days, "label": label})

    return results


def validate_preferences(pref_days, user_name, mmyy, constraints_list, day_duty_counts, all_person_days):
    """
    Validates preference day integers against all applicable constraints.
    Excludes the user's own existing days from counts to avoid double-counting
    when re-submitting.

    Returns: list of str error messages. Empty = no violations.
    """
    from datetime import date as _date
    from collections import defaultdict

    errors = []
    if not pref_days or not constraints_list:
        return errors

    mm = int(mmyy[:2])
    yy = 2000 + int(mmyy[2:])

    existing_person_days = set(all_person_days.get(user_name, []))
    combined_days = sorted(existing_person_days | set(pref_days))

    def adjusted_day_count(day):
        """Day count excluding this person's own existing entry."""
        base = day_duty_counts.get(day, 0)
        return base - (1 if day in existing_person_days else 0)

    for c in constraints_list:
        kind  = c["kind"]
        label = c.get("label", kind)

        if kind == "day_max":
            limit = c["limit"]
            for d in pref_days:
                current = adjusted_day_count(d)
                if current >= limit:
                    errors.append(
                        f"❌ **Day {d}**: already has {current}/{limit} duty slot(s) filled "
                        f"— cannot add another ({label})."
                    )

        elif kind == "person_month_max":
            limit = c["limit"]
            if len(combined_days) > limit:
                errors.append(
                    f"❌ **Monthly cap**: requesting {len(combined_days)} duty day(s) but "
                    f"max per person is {limit}/month ({label})."
                )

        elif kind == "person_week_max":
            limit = c["limit"]
            week_days = defaultdict(list)
            for d in combined_days:
                iso_week = _date(yy, mm, d).isocalendar()[1]
                week_days[iso_week].append(d)
            for week, days_in_week in week_days.items():
                if len(days_in_week) > limit:
                    errors.append(
                        f"❌ **Weekly cap (week {week})**: days {sorted(days_in_week)} gives "
                        f"{len(days_in_week)} duties — max is {limit}/week ({label})."
                    )

        elif kind == "gap_dd":
            gap = c["days"]
            sorted_days = sorted(combined_days)
            for i in range(len(sorted_days) - 1):
                d1, d2 = sorted_days[i], sorted_days[i + 1]
                diff = (_date(yy, mm, d2) - _date(yy, mm, d1)).days
                if diff < gap and (d1 in pref_days or d2 in pref_days):
                    errors.append(
                        f"❌ **Gap violation**: day {d1} and day {d2} are only {diff} day(s) apart "
                        f"— minimum required is {gap} days ({label})."
                    )

    return errors


def get_holiday_duty_days(client, spreadsheet_name, mmyy, user_name):
    """
    Reads the Holiday sheet and returns a list of dicts for any holiday where
    the user (matched case-insensitively against col D or E) is assigned,
    AND the holiday falls in the given mmyy month.
    Each dict: { "date": date_obj, "name": holiday_name_str }
    """
    results = []
    try:
        sh = client.open(spreadsheet_name)
        all_titles = [s.title for s in sh.worksheets()]
        if "Holiday" not in all_titles:
            return results
        hol_ws = sh.worksheet("Holiday")
        hol_raw = hol_ws.get_all_values()

        mm = int(mmyy[:2])
        yy = 2000 + int(mmyy[2:])

        for row in hol_raw[1:]:  # skip header
            if not row or not row[0].strip():
                continue
            hol_holiday_name = row[0].strip()
            hol_date_str = row[1].strip() if len(row) > 1 else ""
            name1 = row[3].strip() if len(row) > 3 else ""
            name2 = row[4].strip() if len(row) > 4 else ""

            # check if this user is assigned
            if user_name.upper() not in (name1.upper(), name2.upper()):
                continue

            # parse date
            hol_date = None
            from datetime import datetime as _dt
            for fmt in ["%d %b %Y", "%-d %b %Y", "%Y-%m-%d"]:
                try:
                    hol_date = _dt.strptime(hol_date_str, fmt).date()
                    break
                except:
                    continue
            if not hol_date:
                continue

            # only include if it falls in the target month
            if hol_date.year == yy and hol_date.month == mm:
                results.append({"date": hol_date, "name": hol_holiday_name})

    except Exception:
        pass
    return results


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