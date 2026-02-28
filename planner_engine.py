import pandas as pd
import numpy as np
import calendar
import time
from datetime import datetime
from collections import defaultdict, Counter
from ortools.sat.python import cp_model
import gspread

def create_backup_and_output(client, spreadsheet_name, mmyy, planned_df, norm_scale, ranges):
    sh = client.open(spreadsheet_name)
    source_name = f"{mmyy}C"
    output_name = f"{mmyy}D"
    
    # duplicate sheet (preserves formatting/validations)
    try:
        sh.del_worksheet(sh.worksheet(output_name))
    except: pass
    
    all_sheets = sh.worksheets()
    last_index = len(all_sheets)

    source_ws = sh.worksheet(source_name)
    output_ws = sh.duplicate_sheet(source_ws.id, insert_sheet_index = last_index, new_sheet_name=output_name)
    
    # prepare updates
    row_start, row_end = ranges['row_start'], ranges['row_end']
    date_start_col, date_end_col = ranges['date_start_col'], ranges['date_end_col']
    offset_col_idx = ranges['constraints_col'] + 1
    
    # clear the duty grid area
    output_ws.batch_clear([f"F{row_start+3}:AJ{row_end+3}"])
    
    updates = []
    for r_idx in range(row_start, row_end + 1):
        gs_row = r_idx + 2
        
        # write normalised points
        pts_val = planned_df.iat[r_idx, offset_col_idx]
        updates.append({'range': gspread.utils.rowcol_to_a1(gs_row, offset_col_idx + 1), 
                        'values': [[round(pts_val, 2)]]})
        
        # writ estimated duties
        est_val = planned_df.loc[r_idx, "Est_Next_Month_Duties"]
        updates.append({'range': gspread.utils.rowcol_to_a1(gs_row, offset_col_idx + 3), 
                        'values': [[est_val]]})

        # write D and S assignments
        for c_idx in range(date_start_col, date_end_col + 1):
            val = planned_df.iat[r_idx, c_idx]
            if val in ["D", "S"]:
                updates.append({'range': gspread.utils.rowcol_to_a1(gs_row, c_idx + 1), 
                                'values': [[val]]})

    # write normal scale to reference cell (AR82)
    updates.append({'range': 'AR82', 'values': [[round(norm_scale, 4)]]})
    
    output_ws.batch_update(updates)
    return output_name

def archive_source_sheet(client, spreadsheet_name, mmyy, folder_id, personal_drive, *args, **kwargs):
    sh = client.open(spreadsheet_name)
    archive_filename = f"[ARCHIVE] {spreadsheet_name}"

    personal_drive.files().copy(
        fileId=sh.id,
        body={
            'name': archive_filename,
            'parents': [folder_id]
        }
    ).execute()

    return archive_filename

def generate_next_month_template(client, spreadsheet_name, mmyy, planned_df, ranges):
    """
    Creates the template for the next month, formats the header,
    and hides columns for days that do not exist in that month.
    """
    sh = client.open(spreadsheet_name)
    
    # calculate next month MMYY
    curr_dt = datetime.strptime(mmyy, "%m%y")
    next_dt = (curr_dt.replace(day=28) + pd.Timedelta(days=4)).replace(day=1)
    next_mmyy = next_dt.strftime("%m%y")
    next_name = f"{next_dt.strftime('%m%y')}C"
    next_spreadsheet_name = f"Plan_Duty_{next_mmyy}"
    
    # duplicate current 'C' sheet to the end
    try:
        sh.del_worksheet(sh.worksheet(next_name))
    except: pass

    all_sheets = sh.worksheets()
    last_index = len(all_sheets)
    
    source_ws = sh.worksheet(f"{mmyy}C")
    next_ws = sh.duplicate_sheet(source_ws.id, insert_sheet_index=last_index, new_sheet_name=next_name)
    
    # setup ranges
    row_start, row_end = ranges['row_start'], ranges['row_end']
    date_start_col = ranges['date_start_col']
    offset_col_idx = ranges['constraints_col'] + 1
    num_days_next = calendar.monthrange(next_dt.year, next_dt.month)[1]

    # format header date
    next_ws.update_acell('F3', next_dt.strftime("%Y-%m-%d"))
    next_ws.format("F3", {
        "numberFormat": {"type": "DATE", "pattern": "dd"},
        "horizontalAlignment": "LEFT",
        "textFormat": {"bold": True}
    })

    # hide unused columns
    # day 1 is index 5 (F), day 31 is index 35 (AJ)
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": next_ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": date_start_col, # column F
                        "endIndex": date_start_col + 31 # up to AJ
                    },
                    "properties": {"hiddenByUser": False},
                    "fields": "hiddenByUser"
                }
            }
        ]
    }
    
    # if month has fewer than 31 days, hide the trailing columns
    if num_days_next < 31:
        body["requests"].append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": next_ws.id,
                    "dimension": "COLUMNS",
                    "startIndex": date_start_col + num_days_next,
                    "endIndex": date_start_col + 31
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser"
            }
        })
    
    sh.batch_update(body)

    # reset grid values and carry points
    updates = []
    for r_idx in range(row_start, row_end + 1):
        gs_row = r_idx + 2
        
        # carry points
        pts_val = planned_df.iat[r_idx, offset_col_idx]
        updates.append({
            'range': gspread.utils.rowcol_to_a1(gs_row, offset_col_idx + 1), 
            'values': [[round(pts_val, 2)]]
        })
        
        # reset grid data
        grid_row = ["" for d in range(1, 32)]
        start_a1 = gspread.utils.rowcol_to_a1(gs_row, date_start_col + 1)
        end_a1 = gspread.utils.rowcol_to_a1(gs_row, date_start_col + 31)
        updates.append({'range': f"{start_a1}:{end_a1}", 'values': [grid_row]})
    
        # for status column
        updates.append({
            'range': f"AS{gs_row}",
            'values': [[""]]
        })
        
    next_ws.batch_update(updates)
    return next_name, next_spreadsheet_name

def run_optimisation(data_bundle, config, point_allocations, model_constraints):
    # --------------------------------------------------
    # SETUP AND DATA EXTRACTION
    # --------------------------------------------------
    constraint_df = data_bundle["constraints"]
    holiday_df = data_bundle["holidays"]
    partners_df = data_bundle["partners"]
    namelist_df = data_bundle["namelist"]
    last_month_df = data_bundle["last_month"]
    year = data_bundle["year"]
    month = data_bundle["month"]
    year_old = data_bundle["year_old"]
    month_old = data_bundle["month_old"]
    
    # Sliders from Web UI
    S1 = config.get('S1', 100)  # partners
    S2 = config.get('S2', 60)  # branch
    S3 = config.get('S3', 10)  # drivers
    S4 = config.get('S4', 400)  # minimum 1x D
    SCALE = 1000

    weekday_points = point_allocations.get('weekday_points', 1.0)
    friday_points = point_allocations.get('friday_points', 1.5)
    sat_sun_points = point_allocations.get('weekend_points', 2.0)
    holiday_points = point_allocations.get('holiday_points', 2.0)

    hard1 = model_constraints.get('hard1', 2)
    hard4 = model_constraints.get('hard4', 4)
    hard5 = model_constraints.get('hard5', 3)
    hard1s = model_constraints.get('hard1s', 2)
    hard2s = model_constraints.get('hard2s', 2)
    scalefactor = model_constraints.get('scalefactor', 4)
    sbf_val = model_constraints.get('sbf_val', 2)

    # --------------------------------------------------
    # INSPECT RANGES
    # --------------------------------------------------
    
    # determine row range (people)
    name_col_idx = 1  # "name" column

    name_series = constraint_df.iloc[:, name_col_idx]

    # find first empty cell after header row (row 1 onwards)
    first_empty_row = None
    for i in range(1, len(name_series)):
        val = name_series.iloc[i]
        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
            first_empty_row = i
            break

    row_start = 2
    row_end = first_empty_row - 1

    # determine column range (dates)

    date_start_col = 5

    num_days = calendar.monthrange(year, month)[1]

    date_end_col = date_start_col + num_days - 1

    # determine column range (constraints)

    constraints_col = 42 # column AQ
    OFFSET_COL = constraints_col + 1

    # --------------------------------------------------
    # AVAILABILITY_DF
    # --------------------------------------------------

    availability_df = constraint_df.copy()

    for r in range(row_start, row_end + 1):
        for q in range(date_start_col, date_end_col + 1):
            cell = availability_df.iat[r, q]
            if pd.isna(cell):
                availability_df.iat[r, q] = 1
            else:
                availability_df.iat[r, q] = 0


    # --------------------------------------------------
    # HOLIDAY POINTS
    # --------------------------------------------------
    
    # ensure the DATE column is datetime
    holiday_df['DATE'] = pd.to_datetime(holiday_df['DATE'], errors='coerce')

    # filter by year and month
    mask = (holiday_df['DATE'].dt.year == year) & (holiday_df['DATE'].dt.month == month)
    holiday_filtered = holiday_df[mask]

    # generate all days in the month
    num_days = pd.Period(f'{year}-{month:02d}').days_in_month
    dates = pd.date_range(start=f'{year}-{month:02d}-01', periods=num_days)

    # construct the dataFrame
    points_df = pd.DataFrame({
        'date': dates.day,
        'day': dates.day_name().str[:3].str.lower(),
        'iso': dates.isocalendar().week.astype(int),
        'points': weekday_points  # default for Mon-Thu
    })

    # assign points for Fri/Sat/Sun
    points_df.loc[points_df['day'] == 'fri', 'points'] = friday_points
    points_df.loc[points_df['day'].isin(['sat', 'sun']), 'points'] = sat_sun_points

    # override points for holidays
    # convert holiday_filtered['DATE'] to day numbers in the month
    holiday_days = holiday_filtered['DATE'].dt.day.tolist()
    points_df.loc[points_df['date'].isin(holiday_days), 'points'] = holiday_points

    date_points = {
        c: int(points_df.iloc[i]["points"] * SCALE)
        for i,c in enumerate(range(date_start_col, date_end_col+1))
    }

    holiday_days = holiday_filtered['DATE'].dt.day.tolist()
    holiday_cols = {date_start_col + (d - 1) for d in holiday_days}

    fix_assignment_df = availability_df.copy()

    # --------------------------------------------------
    # STAFF ATTRIBUTES AND HISTORY
    # --------------------------------------------------
    name_to_row = {}

    # indentify female staff
    is_female_pair = {}
    female_indices = []

    for r in range(row_start, row_end + 1):
        name = str(constraint_df.iat[r, 1]).strip().upper()
        
        if "(F)" in name:
            is_female_pair[r] = True
            female_indices.append(r)
        else:
            is_female_pair[r] = False

    # last month weekend and duties
    weekend_workers_last_month = set()
    duties_last_month = defaultdict(list)
    if isinstance(last_month_df, pd.DataFrame):
        _, last_num_days = calendar.monthrange(year_old, month_old)

        for c in range(date_start_col, date_start_col + last_num_days):
            day_num = c - date_start_col + 1
            last_date = datetime(year_old, month_old, day_num)

            for r in range(row_start, row_end + 1):
                cell_val = str(last_month_df.iat[r, c]).strip().upper()
                if cell_val == "D":
                    name = str(last_month_df.iat[r, 1]).strip().upper()
                    duties_last_month[name].append(day_num)

                    if last_date.weekday() in [5, 6]:
                        weekend_workers_last_month.add(name)

    # --------------------------------------------------
    # SOFT CONSTRAINT SETUP
    # --------------------------------------------------
    
    # soft constraint 1 set-up

    name_to_row = {}
    for r in range(row_start, row_end + 1):
        name = str(constraint_df.iat[r, 1]).strip().upper()
        name_to_row[name] = r

    partner_pairs = []
    for i in range(1, len(partners_df)):
        p1_name = str(partners_df.iloc[i, 1]).strip().upper()
        p2_name = str(partners_df.iloc[i, 3]).strip().upper()

        if p1_name in name_to_row and p2_name in name_to_row:
            partner_pairs.append((name_to_row[p1_name], name_to_row[p2_name]))

    # soft constraint 2 set-up

    b_name_to_row = {}

    for r in range(row_start, row_end + 1):
        raw_name = constraint_df.iat[r, 1]

        if pd.isna(raw_name):
            continue

        name = str(raw_name).strip().upper()

        b_name_to_row[name] = r

    branch_to_row = defaultdict(list)
    unmatched_names = []

    for i in range(len(namelist_df)):
        raw_name = namelist_df.iloc[i, 1]
        raw_branch = namelist_df.iloc[i, 2]

        if pd.isna(raw_name) or pd.isna(raw_branch):
            continue

        name = str(raw_name).strip().upper()
        branch = str(raw_branch).strip().upper()

        if name in b_name_to_row:
            duty_row = b_name_to_row[name]
            branch_to_row[branch].append(duty_row)
        else:
            unmatched_names.append(name)

    # soft constraint 3 set-up

    is_driver = {}
    is_non_driver = {}

    for i in range(1, len(partners_df)):

        # column 1 & 2 (pair 1), column 3 & 4 (pair 2)
        pairs = [(1, 2), (3, 4)]

        for name_col, desig_col in pairs:
            person_name = str(partners_df.iloc[i, name_col]).strip().upper()
            designation = str(partners_df.iloc[i, desig_col]).strip().upper()

            # skip processing if the name cell is empty or invalid
            if person_name == "NAN" or not person_name:
                continue

            if person_name in name_to_row:
                row_idx = name_to_row[person_name]

                # categorise based on the designation text
                if designation == "DRIVER":
                    is_driver[row_idx] = True
                else:
                    # any other designation defaults to non-driver
                    is_non_driver[row_idx] = True
    
    # partner set-up

    row_to_partner = {}
    for r1, r2 in partner_pairs:
        row_to_partner[r1] = r2
        row_to_partner[r2] = r1

    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()

        if "SBF" in status_val:
            # find this person's partner using the lookup
            partner_row_idx = row_to_partner.get(r)

            if partner_row_idx is not None:
                # check the PARTNER'S current status
                current_partner_status = str(fix_assignment_df.iat[partner_row_idx, OFFSET_COL + 1]).strip().upper()

                # only label as 'PARTNER' if they aren't already 'SBF' themselves
                if "SBF" not in current_partner_status:
                    partner_name = fix_assignment_df.iat[partner_row_idx, 1]
                    fix_assignment_df.iat[partner_row_idx, OFFSET_COL + 1] = "PARTNER"


    # --------------------------------------------------
    # HARD OBJECTIVES
    # --------------------------------------------------
    col_to_date = {c: datetime(year, month, c - date_start_col + 1)
                for c in range(date_start_col, date_end_col + 1)}

    iso_map = defaultdict(list)
    for c, dt in col_to_date.items():
        iso_week = dt.isocalendar().week
        iso_map[iso_week].append(c)

    model = cp_model.CpModel()
    x = {}
    fixed_duties = set()
    fixed_standbys = set()

    exclusion_keywords = ["SBF", "SAIL", "NDP", "EXCUSED", "MEDICAL", "ON COURSE", "PARTNER"]

    # variable creation and initial constraints
    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        is_excluded_for_month = any(keyword in status_val for keyword in exclusion_keywords)

        for c in range(date_start_col, date_end_col + 1):
            cell = str(constraint_df.iat[r, c]).strip().upper() if not pd.isna(constraint_df.iat[r, c]) else ""
            is_holiday = c in holiday_cols
            x[(r, c)] = model.NewBoolVar(f"x_{r}_{c}")

            if is_excluded_for_month:
                if is_holiday and cell == "D":
                    model.Add(x[(r, c)] == 1)
                    fixed_duties.add((r, c))
                else:
                    model.Add(x[(r, c)] == 0)
            else:
                if cell == "D":
                    model.Add(x[(r, c)] == 1)
                    fixed_duties.add((r, c))
                elif cell == "S":
                    model.Add(x[(r, c)] == 0)
                    fixed_standbys.add((r, c))
                elif cell == "X":
                    model.Add(x[(r, c)] == 0)

    # hard constraint 1: must have hard1 Ds a day
    for c in range(date_start_col, date_end_col + 1):
        model.Add(sum(x[(r, c)] for r in range(row_start, row_end + 1)) == hard1)

    # hard constraint 2: no weekend duty if worked weekend last month (unless manually assigned D)
    for r in range(row_start, row_end + 1):
        staff_name = str(fix_assignment_df.iat[r, 1]).strip().upper()

        if staff_name in weekend_workers_last_month:
            for c in range(date_start_col, date_end_col + 1):
                # skip if day was manually assigned a D
                if (r, c) in fixed_duties:
                    continue

                # check if current day is a weekend
                current_day_num = c - date_start_col + 1
                current_date = datetime(year, month, current_day_num)

                if current_date.weekday() in [5, 6]:
                    if (r, c) in x:
                        model.Add(x[(r, c)] == 0)

    # hard constraint 3: 1x D per week (unless manually assigned D)
    for r in range(row_start, row_end + 1):
        for week, cols in iso_map.items():
            vars_in_week = [x[(r, c)] for c in cols if (r, c) not in fixed_duties or c not in holiday_cols]
            if vars_in_week:
                model.Add(sum(vars_in_week) <= 1)

    # hard constraint 4: cross-month and internal gap must be at least hard4 days
    for r in range(row_start, row_end + 1):
        staff_name = str(fix_assignment_df.iat[r, 1]).strip().upper()
        last_month_list = duties_last_month.get(staff_name, [])

        # cross-month gap
        if last_month_list:
            # get the very last duty day from the previous month
            last_d_num = max(last_month_list)
            last_d_date = datetime(year_old, month_old, last_d_num)

            for c in range(date_start_col, date_end_col + 1):
                # skip if there is a manual D
                if (r, c) in fixed_duties:
                    continue

                current_date = col_to_date[c]
                # calculate gap
                if (current_date - last_d_date).days < hard4:
                    model.Add(x[(r, c)] == 0)
                else:
                    break

        # internal gap
        for c1 in range(date_start_col, date_end_col + 1):
            d1 = col_to_date[c1]
            for c2 in range(c1 + 1, date_end_col + 1):
                # skip if there is a manual D
                if (r, c1) in fixed_duties or (r, c2) in fixed_duties:
                    continue

                # 4-day gap
                if (col_to_date[c2] - d1).days >= hard4:
                    break
                model.Add(x[(r, c1)] + x[(r, c2)] <= 1)

    # hard constraint 5: keep D count to <= hard5
    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        is_excluded = any(k in status_val for k in exclusion_keywords)

        total_duties_vars = [x[(r, c)] for c in range(date_start_col, date_end_col + 1)]

        if is_excluded:
            num_holiday_duties = sum(1 for c in holiday_cols if (r, c) in fixed_duties)
            model.Add(sum(total_duties_vars) == num_holiday_duties)
        elif is_female_pair.get(r, False):
            # females are exempted
            model.Add(sum(total_duties_vars) <= hard5 - 1)
        else:
            model.Add(sum(total_duties_vars) <= hard5)

    # hard constraint 6: females must do duty together

    for c in range(date_start_col, date_end_col + 1):
        female_vars_on_day = [x[(r, c)] for r in female_indices if (r, c) in x]
        
        if female_vars_on_day:
            # Create a helper variable to represent the count of females
            female_count = model.NewIntVar(0, 2, f"female_count_day_{c}")
            model.Add(female_count == sum(female_vars_on_day))
            
            # This forces the count to be either 0 or 2 (never 1)
            is_female = model.NewBoolVar(f"is_female_{c}")
            model.Add(female_count == 2).OnlyEnforceIf(is_female)
            model.Add(female_count == 0).OnlyEnforceIf(is_female.Not())

    # --------------------------------------------------
    # SOFT CONSTRAINTS
    # --------------------------------------------------

    soft_penalties = []

    # soft constraint 1: preferred pairings

    for r1, r2 in partner_pairs:
        for c in range(date_start_col, date_end_col + 1):
            if (r1, c) in x and (r2, c) in x:
                # create a boolean variable: 1 if they are split, 0 if they are together
                is_split = model.NewBoolVar(f"split_{r1}_{r2}_day_{c}")
                # Logic: is_split must be 1 if x[r1] != x[r2]
                model.Add(x[(r1, c)] - x[(r2, c)] <= is_split)
                model.Add(x[(r2, c)] - x[(r1, c)] <= is_split)

                soft_penalties.append(is_split * S1)

    # soft constraint 2: different branches

    for c in range(date_start_col, date_end_col + 1):
        for branch, row_indices in branch_to_row.items():
            vars_in_branch = [x[(r,c)] for r in row_indices if (r, c) in x]

            if len(vars_in_branch) >= 2:
                same_branch_violation = model.NewBoolVar(f"branch_violation_{branch}_{c}")

                model.Add(sum(vars_in_branch) < 2).OnlyEnforceIf(same_branch_violation.Not())
                model.Add(sum(vars_in_branch) == 2).OnlyEnforceIf(same_branch_violation)

                soft_penalties.append(same_branch_violation * S2)

    # soft constraint 3: driver/non-driver/rider combination

    for c in range(date_start_col, date_end_col + 1):
        drivers_on_day = [x[(r,c)] for r in range(row_start, row_end + 1) if (r, c) in x and r in is_driver]

    if drivers_on_day:
        driver_count = model.NewIntVar(0, 2, f"driver_count_day_{c}")
        model.Add(driver_count == sum(drivers_on_day))

        mismatch = model.NewBoolVar(f"driver_mismatch_day_{c}")

        model.Add(driver_count != 1).OnlyEnforceIf(mismatch)
        model.Add(driver_count == 1).OnlyEnforceIf(mismatch.Not())

        soft_penalties.append(mismatch * S3)

    # soft constraint 4: minimum 1 day

    has_at_least_one_duty = {}
    for r in range(row_start, row_end + 1):
        # create a boolean variable: 1 if staff has >= 1 duty, 0 otherwise
        has_at_least_one_duty[r] = model.NewBoolVar(f'has_duty_{r}')

        # sum of duties for this specific staff member
        duties_sum = sum(x[(r, c)] for c in range(date_start_col, date_end_col + 1) if (r, c) in x)

        # if has_at_least_one_duty is True (1), then duties_sum must be >= 1
        # if has_at_least_one_duty is False (0), then duties_sum must be 0
        model.Add(duties_sum >= 1).OnlyEnforceIf(has_at_least_one_duty[r])
        model.Add(duties_sum == 0).OnlyEnforceIf(has_at_least_one_duty[r].Not())

        # penalize the NEGATION of having a duty
        soft_penalties.append(has_at_least_one_duty[r].Not() * S4)

    # --------------------------------------------------
    # FAIRNESS OBJECTIVE
    # --------------------------------------------------

    # define point scales and constants
    SBF_BONUS = int(sbf_val * SCALE)

    final_scores = {}
    duty_counts = {}

    sum_new_points = 0
    sum_new_points = sum(
        int(round(p * SCALE * 2))
        for p in points_df["points"]
    )

    try:
        # row 81 and col 43 because pandas is 0-indexed
        last_month_scale = float(last_month_df.iat[81, 43])
        if last_month_scale <= 0: last_month_scale = 1.0
    except:
        last_month_scale = 1.0 # Default fallback

    # adjust for any changes to the sheet last month

    manual_adjustments = {}
    if isinstance(last_month_df, pd.DataFrame):
        # iterate from row 36 onwards
        for i in range(36, row_end+ 1):
            raw_name = last_month_df.iat[i, 48]
            if pd.isna(raw_name) or str(raw_name).strip() == "":
                continue
            name = str(raw_name).strip().upper()
            change_text = str(last_month_df.iat[i, 49]).strip().upper() # column AX
            base_val = last_month_df.iat[i, 50] # column AY value

            if not name or name == "NAN": continue

            adj = 0
            if "ADD 1X WD" in change_text:    adj = (1 / last_month_scale)
            elif "ADD 1X F" in change_text:   adj = (1.5 / last_month_scale)
            elif "ADD 1X WE" in change_text:  adj = (2 / last_month_scale)
            elif "ADD 1X H" in change_text: adj = (2 / last_month_scale)
            elif "MINUS 1X WD" in change_text:    adj = -(1 / last_month_scale)
            elif "MINUS 1X F" in change_text:   adj = -(1.5 / last_month_scale)
            elif "MINUS 1X WE" in change_text:  adj = -(2 / last_month_scale)
            elif "MINUS 1X H" in change_text: adj = -(2 / last_month_scale)

            manual_adjustments[name] = base_val + adj

    # calculate scores and counts loop
    for r in range(row_start, row_end + 1):

        # pull changes from last month
        staff_name = str(fix_assignment_df.iat[r,1]).strip().upper()

        if staff_name in manual_adjustments:
            val = manual_adjustments[staff_name]
        else:
            val = fix_assignment_df.iat[r, OFFSET_COL]

        try:
            # convert to float first to handle strings that look like numbers
            numeric_val = float(val) if (val is not None and str(val).strip() != "") else 0.0
        except ValueError:
            # if it's actual text like "SBF", default to 0.0
            numeric_val = 0.0

        current_scaled = int(round(numeric_val * SCALE))

        # SBF and NEW points
        status_cell = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        bonus_points = 0

        if "SBF" in status_cell:
            bonus_points += SBF_BONUS

        if "NEW" in status_cell:
            mask = last_month_df.astype(str).apply(lambda col: col.str.contains("Avg Offset", case=False, na=False))
            matches = np.where(mask.values)
            if len(matches[0]) > 0:
                row_idx, col_idx = matches[0][0], matches[1][0]
                avg_val = last_month_df.iat[row_idx, col_idx + 1]
                bonus_points += int(round(avg_val * SCALE))

        # current month points
        assigned_points_expr = sum(
            x[(r, c)] * int(round(points_df.iloc[c - date_start_col]["points"] * SCALE))
            for c in range(date_start_col, date_end_col + 1)
            if (r, c) in x
        )

        # total score variable
        total_score = model.NewIntVar(0, sum_new_points + (100 * SCALE), f"total_score_{r}")
        model.Add(total_score == assigned_points_expr + current_scaled + bonus_points)
        final_scores[r] = total_score

        # duty count variable
        num_duties = sum(x[(r, c)] for c in range(date_start_col, date_end_col + 1) if (r, c) in x)
        count_var = model.NewIntVar(0, 31, f"duty_count_{r}")
        model.Add(count_var == num_duties)
        duty_counts[r] = count_var

    # min/max logic
    max_score = model.NewIntVar(0, sum_new_points + (100 * SCALE), "max_score")
    min_score = model.NewIntVar(0, sum_new_points + (100 * SCALE), "min_score")
    max_duties = model.NewIntVar(0, 31, "max_duties")
    min_duties = model.NewIntVar(0, 31, "min_duties")

    for r in range(row_start, row_end + 1):
        model.Add(max_score >= final_scores[r])
        model.Add(min_score <= final_scores[r])
        model.Add(max_duties >= duty_counts[r])
        model.Add(min_duties <= duty_counts[r])

    # define gaps
    score_gap = model.NewIntVar(0, sum_new_points + (100 * SCALE), "score_gap")
    model.Add(score_gap == max_score - min_score)

    duty_day_gap = model.NewIntVar(0, 31, "duty_day_gap")
    model.Add(duty_day_gap == max_duties - min_duties)

    # final objective
    model.Minimize(
        score_gap +
        (duty_day_gap * 100) +
        sum(soft_penalties)
    )

    # --------------------------------------------------
    # SOLVER
    # --------------------------------------------------

    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = 8
    solver.parameters.max_time_in_seconds = 15

    start_time = time.time()
    status = solver.Solve(model)
    end_time = time.time()

    current_time = datetime.now()
    print("\nCode ran at ", current_time)
    print(f"Solver finished in {end_time - start_time:.2f} seconds")

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for r, score_var in final_scores.items():
            staff_name = fix_assignment_df.iat[r, 1]
            actual_score = solver.Value(score_var)
    else:
        print("No feasible solution found")

    # --------------------------------------------------
    # OFFSET CALCULATION
    # --------------------------------------------------

    results = []

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for r in range(row_start, row_end + 1):
            staff_name = fix_assignment_df.iat[r, 1]

            # scaled points: start with current points
            current_scaled = int(round(fix_assignment_df.iat[r, OFFSET_COL] * SCALE)) \
                if not pd.isna(fix_assignment_df.iat[r, OFFSET_COL]) else 0

            # total points including model-assigned duties
            total_scaled = solver.Value(final_scores[r]) if r in final_scores else current_scaled

            # scale back down
            total_points = total_scaled / SCALE

            # now iterate through dates to get duties
            for c in range(date_start_col, date_end_col + 1):
                if (r, c) in x and solver.Value(x[(r, c)]) == 1:
                    duty_date = fix_assignment_df.iat[0, c]
                    points = points_df.iloc[c - date_start_col]["points"]  # already raw points
                    fixed = (r, c) in fixed_duties
                    results.append({
                        "Name": staff_name,
                        "Date": duty_date,
                        "Points": points,
                        "Fixed": fixed,
                        "Updated_Total_Points": total_points
                    })

    results_df = pd.DataFrame(results) # for raw data output

    planned_df = fix_assignment_df.copy()

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for (r, c), var in x.items():
            if r in final_scores:
                planned_df.iat[r, OFFSET_COL] = solver.Value(final_scores[r]) / SCALE

            if solver.Value(var) == 1:
                if planned_df.iat[r, c] != "D":
                    planned_df.iat[r, c] = "D"

    # normalisation of points
    active_rows = [
        r for r in range(row_start, row_end + 1)
    ]

    points_list = [planned_df.iat[r, OFFSET_COL] for r in active_rows]

    min_points = min(points_list)
    max_points = max(points_list)
    diff_points = max_points - min_points

    # scaling of offset
    normal_scale = 1
    for r in range(row_start, row_end + 1):
        if max_points != min_points:

            # if disparity less than scalefactor, no change
            if diff_points <= scalefactor: # arbitrary value
                planned_df.iat[r, OFFSET_COL] = (planned_df.iat[r, OFFSET_COL] - min_points)

            # if disparity more than scalefactor, scale down
            else:
                planned_df.iat[r, OFFSET_COL] = (planned_df.iat[r, OFFSET_COL] - min_points)
                normal_scale = diff_points / scalefactor
                planned_df.iat[r, OFFSET_COL] = planned_df.iat[r, OFFSET_COL] / normal_scale
        else:
            planned_df.iat[r, OFFSET_COL] = 0  # all equal
            normal_scale = 1

    # --------------------------------------------------
    # NEXT MONTH'S PROJECTION
    # --------------------------------------------------

    month_new = month + 1 if month < 12 else 1
    year_new = year if month < 12 else year + 1
    
    num_days_new = calendar.monthrange(year_new, month_new)[1]
    dates_new = pd.date_range(start=f'{year_new}-{month_new:02d}-01', periods=num_days_new)

    # use holiday_df to find holidays in the projected month
    holiday_df['DATE'] = pd.to_datetime(holiday_df['DATE'], errors='coerce')
    mask_new = (holiday_df['DATE'].dt.year == year_new) & (holiday_df['DATE'].dt.month == month_new)
    holiday_days_new = holiday_df[mask_new]['DATE'].dt.day.tolist()

    total_points_next_month = 0
    for d in dates_new:
        day_name = d.day_name()[:3].lower()
        # holiday points
        p = 2.0 if (d.day in holiday_days_new or day_name in ['sat', 'sun']) else (1.5 if day_name == 'fri' else 1.0)
        total_points_next_month += (p * 2) # Two people assigned per day

    # distribute based on current offset (Higher points today = Lower priority tomorrow)
    current_offsets = [planned_df.iat[r, OFFSET_COL] for r in range(row_start, row_end + 1)]
    max_off = max(current_offsets)
    weights = [(max_off + 1) - off for off in current_offsets]
    total_weight = sum(weights)

    for i, r in enumerate(range(row_start, row_end + 1)):
        share = weights[i] / total_weight if total_weight > 0 else (1 / len(weights))
        # Estimated days based on avg points (1.4 factor)
        est_days = (share * total_points_next_month) / 1.4
        planned_df.loc[r, "Est_Next_Month_Duties"] = round(est_days, 1)

    # --------------------------------------------------
    # STANDBY
    # --------------------------------------------------

    model_s = cp_model.CpModel()
    s = {}

    # availability creation

    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        is_excluded = any(k in status_val for k in exclusion_keywords)

        for c in range(date_start_col, date_end_col + 1):
            # skip if the person is generally excluded for the month
            if is_excluded or is_female_pair.get(r, False):
                continue

            # skip if the person is already assigned a D from Pass 1
            if planned_df.iat[r, c] == "D":
                continue

            # skip if the cell is marked X
            cell_val = str(constraint_df.iat[r, c]).strip().upper() if not pd.isna(constraint_df.iat[r, c]) else ""
            if cell_val == "X":
                continue

            # otherwise, create the S variable
            s[(r, c)] = model_s.NewBoolVar(f"s_{r}_{c}")

    # hard constraint 1S: must have hard1s S per day

    for c in range(date_start_col, date_end_col + 1):
        day_vars = [s[(r, c)] for r in range(row_start, row_end + 1) if (r, c) in s]
        if day_vars:
            model_s.Add(sum(day_vars) == hard1s)

    # hard constraint 2S: must have hard2s days between D and S

    for r in range(row_start, row_end + 1):

        for c1 in range(date_start_col, date_end_col + 1):
            d1 = col_to_date[c1]

            for c2 in range(c1 + 1, date_end_col + 1):
                d2 = col_to_date[c2]
                day_diff = (d2 - d1).days

                if day_diff < hard2s:

                    # check D at c1, S at c2
                    if planned_df.iat[r, c1] == "D" and (r, c2) in s:
                        model_s.Add(s[(r, c2)] == 0)

                    # check S at c1, D at c2
                    if (r, c1) in s and planned_df.iat[r, c2] == "D":
                        model_s.Add(s[(r, c1)] == 0)

                    # check between S and S
                    if (r, c1) in s and (r, c2) in s:
                        model_s.Add(s[(r, c1)] + s[(r, c2)] <= hard2s - 1)

                else:
                    break

    # hard constraint 3S: S must be in same branch as D

    from collections import Counter

    for c in range(date_start_col, date_end_col + 1):

        # identify D per day
        d_row_indices = [
            r for r in range(row_start, row_end + 1)
            if planned_df.iat[r, c] == "D"
        ]

        num_d = len(d_row_indices)

        # if there is no D in the branch, no S allowed
        day_s_vars = [s[(r, c)] for r in range(row_start, row_end + 1) if (r, c) in s]

        if num_d == 0:
            for var in day_s_vars:
                model_s.Add(var == 0)
            continue

        # map D to branches
        required_branches = []
        for r_idx in d_row_indices:
            for branch, rows in branch_to_row.items():
                if r_idx in rows:
                    required_branches.append(branch)
                    break

        branch_counts = Counter(required_branches)

        # total S must equal to total D
        if len(day_s_vars) < num_d:
            raise ValueError(
                f"Day {c}: {len(day_s_vars)} S candidates for {num_d} Ds"
            )

        model_s.Add(sum(day_s_vars) == num_d)

        # each branch gets the same no. of S as D
        for branch, required_s in branch_counts.items():
            row_indices = branch_to_row[branch]
            branch_s_vars = [s[(r, c)] for r in row_indices if (r, c) in s]

            if len(branch_s_vars) < required_s:
                raise ValueError(
                    f"Day {c}, Branch {branch}: "
                    f"needs {required_s} S but only {len(branch_s_vars)} available"
                )

            model_s.Add(sum(branch_s_vars) == required_s)

        # branches with D have no S
        for branch, row_indices in branch_to_row.items():
            if branch not in branch_counts:
                for r in row_indices:
                    if (r, c) in s:
                        model_s.Add(s[(r, c)] == 0)

    # count the number of S and ensure it doesnt not exceed 5
    s_counts = {}
    for r in range(row_start, row_end + 1):
        personal_s_vars = [s[(r, c)] for c in range(date_start_col, date_end_col + 1) if (r, c) in s]
        count_var = model_s.NewIntVar(0, 5, f"s_count_{r}")
        model_s.Add(count_var == sum(personal_s_vars))
        s_counts[r] = count_var

    # minimize the max standby assignments per person
    max_s = model_s.NewIntVar(0, 5, "max_s")
    for count_var in s_counts.values():
        model_s.Add(max_s >= count_var)
    model_s.Minimize(max_s)

    solver_s = cp_model.CpSolver()
    solver_s.parameters.max_time_in_seconds = 10
    status_s = solver_s.Solve(model_s)

    if status_s in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("Standby pass successful.")
        for (r, c), var in s.items():
            if solver_s.Value(var) == 1:
                planned_df.iat[r, c] = "S"
    else:
        print("Could not find a feasible solution for Standby.")


    return planned_df, normal_scale, {
        'row_start': row_start,
        'row_end': row_end,
        'date_start_col': date_start_col,
        'date_end_col': date_end_col,
        'constraints_col': constraints_col
    }