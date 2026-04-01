import pandas as pd
import numpy as np
import calendar
import time
from datetime import datetime
from collections import defaultdict, Counter
from ortools.sat.python import cp_model
import json
from dynamic_constraints import apply_dynamic_constraints
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
    
    # clear the duty grid area (col E to AI, data starts Excel row 4)
    output_ws.batch_clear([f"E{row_start+4}:AI{row_end+4}"])
    
    updates = []
    for r_idx in range(row_start, row_end + 1):
        gs_row = r_idx + 4  # data starts at Excel row 4
        
        # write normalised points
        pts_val = planned_df.iat[r_idx, offset_col_idx]
        updates.append({'range': 'AQ' + str(gs_row), 'values': [[round(pts_val, 2)]]})
        
        # write estimated duties (AS = offset_col_idx + 2)
        est_val = planned_df.loc[r_idx, "Est_Next_Month_Duties"]
        updates.append({'range': 'AS' + str(gs_row), 'values': [[est_val]]})

        # write D and S assignments
        for c_idx in range(date_start_col, date_end_col + 1):
            val = planned_df.iat[r_idx, c_idx]
            if val in ["D", "S"]:
                updates.append({'range': gspread.utils.rowcol_to_a1(gs_row, c_idx + 1), 
                                'values': [[val]]})

    # write normal scale to reference cell (AR82)
    updates.append({'range': 'AU3', 'values': [[round(norm_scale, 4)]]})
    
    output_ws.batch_update(updates)
    return output_name

def archive_source_sheet(client, spreadsheet_name, mmyy, folder_id, personal_drive, *args, **kwargs):
    sh = client.open(spreadsheet_name)
    archive_filename = f"[ARCHIVE] {spreadsheet_name}"
    temp_filename = f"[ARCHIVE] {spreadsheet_name} 1"

    # step 1: find existing archive and rename it to temp name
    existing_query = (
        f"name = '{archive_filename}' "
        f"and trashed = false "
        f"and '{folder_id}' in parents"
    )
    existing_results = personal_drive.files().list(q=existing_query, fields="files(id)").execute()
    existing_files = existing_results.get('files', [])

    old_archive_id = None
    if existing_files:
        old_archive_id = existing_files[0]['id']
        personal_drive.files().update(
            fileId=old_archive_id,
            body={'name': temp_filename}
        ).execute()

    # step 2: create the new archive as usual
    personal_drive.files().copy(
        fileId=sh.id,
        body={
            'name': archive_filename,
            'parents': [folder_id]
        }
    ).execute()

    # step 3: delete the old archive (now renamed to temp)
    if old_archive_id:
        personal_drive.files().delete(fileId=old_archive_id).execute()

    # step 4: empty the trash
    personal_drive.files().emptyTrash().execute()

    return archive_filename

def generate_next_month_template(client, spreadsheet_name, mmyy, planned_df, ranges):
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
    first_day_str = next_dt.strftime("%Y-%m-%d")
    next_ws.update_acell('E3', first_day_str)
    next_ws.format("E3", {
        "numberFormat": {"type": "DATE", "pattern": "dd"},
        "horizontalAlignment": "LEFT",
        "textFormat": {"bold": True}
    })

    # hide unused columns
    # col E = sheet index 4, col AI = sheet index 34
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": next_ws.id,
                        "dimension": "COLUMNS",
                        "startIndex": 4,   # col E (0-indexed)
                        "endIndex": 35     # col AI inclusive (0-indexed)
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
                    "startIndex": 4 + num_days_next,   # first unused day col
                    "endIndex": 35                      # through AI
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser"
            }
        })
    
    sh.batch_update(body)

    # reset grid values and carry points
    updates = []
    for r_idx in range(row_start, row_end + 1):
        gs_row = r_idx + 4  # data starts at Excel row 4
        
        # carry points (AQ = offset_col_idx + 1 in 1-indexed gspread)
        pts_val = planned_df.iat[r_idx, offset_col_idx]
        updates.append({
            'range': gspread.utils.rowcol_to_a1(gs_row, offset_col_idx + 1), 
            'values': [[round(pts_val, 2)]]
        })
        
        # reset grid data (col E to AI = cols 5 to 35 in 1-indexed gspread)
        grid_row = ["" for _ in range(31)]
        start_a1 = gspread.utils.rowcol_to_a1(gs_row, 5)   # col E
        end_a1 = gspread.utils.rowcol_to_a1(gs_row, 35)    # col AI
        updates.append({'range': f"{start_a1}:{end_a1}", 'values': [grid_row]})
    
        # reset status column (AR = col 44 in 1-indexed gspread)
        updates.append({
            'range': f"AR{gs_row}",
            'values': [[""]]
        })

        # reset forecast column (AS = col 45 in 1-indexed gspread)
        updates.append({
            'range': f"AS{gs_row}",
            'values': [[""]]
        })
    

        
    next_ws.batch_update(updates)
    return next_name, next_spreadsheet_name

def run_optimisation(data_bundle, config, point_allocations, model_constraints, slider_overrides=None):
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
    
    SCALE = 1000

    weekday_points = point_allocations.get('weekday_points', 1.0)
    friday_points = point_allocations.get('friday_points', 1.0)
    sat_sun_points = point_allocations.get('weekend_points', 2.0)
    holiday_points = point_allocations.get('holiday_points', 2.0)

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
    for i in range(0, len(name_series)):
        val = name_series.iloc[i]
        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
            first_empty_row = i
            break

    row_start = 0
    row_end = first_empty_row - 1

    # determine column range (dates)

    date_start_col = 4

    num_days = calendar.monthrange(year, month)[1]

    date_end_col = date_start_col + num_days - 1

    # determine column range (constraints)

    constraints_col = 41 # column AQ
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

    # --------------------------------------------------
    # SOFT CONSTRAINT SETUP
    # --------------------------------------------------
    
    # soft constraint 1 set-up

    name_to_row = {}
    for r in range(row_start, row_end + 1):
        name = str(constraint_df.iat[r, 1]).strip().upper()
        name_to_row[name] = r

    # Partners structure: col B = Names (everyone), col C = Partner
    # each valid pair appears twice — deduplicate with frozensets
    partner_pairs = []
    seen_pairs = set()
    for i in range(0, len(partners_df)):
        p1_name = str(partners_df.iloc[i, 1]).strip().upper()  # Names col = index 1
        p2_name = str(partners_df.iloc[i, 2]).strip().upper()  # Partner col = index 2

        if not p1_name or not p2_name or p2_name in ("NAN", "") or p1_name in ("NAN", ""):
            continue
        pair_key = frozenset([p1_name, p2_name])
        if pair_key in seen_pairs:
            continue
        if p1_name in name_to_row and p2_name in name_to_row:
            partner_pairs.append((name_to_row[p1_name], name_to_row[p2_name]))
            seen_pairs.add(pair_key)

    branch_to_row = defaultdict(list)

    for i in range(len(namelist_df)):
        raw_name = namelist_df.iloc[i, 1]
        raw_branch = namelist_df.iloc[i, 2]

        if pd.isna(raw_name) or pd.isna(raw_branch):
            continue

        name = str(raw_name).strip().upper()
        branch = str(raw_branch).strip().upper()

        if name in name_to_row:
            duty_row = name_to_row[name]
            branch_to_row[branch].append(duty_row)

    # soft constraint 3 set-up

    is_driver = {}

    for i in range(len(namelist_df)):
        person_name = str(namelist_df.iloc[i, 1]).strip().upper()
        designation = str(namelist_df.iloc[i, 3]).strip().upper()  # column D

        if person_name == "NAN" or not person_name:
            continue

        if person_name in name_to_row:
            row_idx = name_to_row[person_name]
            if designation == "DRIVER":
                is_driver[row_idx] = True
    
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

            # 1. PRIORITY: If there is a manual "D", always assign the duty
            # regardless of SBF or other exclusion status
            if cell == "D":
                x[(r, c)] = model.NewConstant(1)
                fixed_duties.add((r, c))
                continue

            # 2. Create the variable for non-manual days
            x[(r, c)] = model.NewBoolVar(f"x_{r}_{c}")

            # 3. Excluded (SBF/excused/partner), S, or X cells all block duty assignment.
            #    S is handled fully in the standby pass below — it just means unavailable here.
            if is_excluded_for_month or cell in ("S", "X"):
                model.Add(x[(r, c)] == 0)

    # --------------------------------------------------
    # DYNAMIC CONSTRAINTS (interpreted from CONFIG sheet)
    # --------------------------------------------------
    soft_penalties, has_at_least_one_duty = apply_dynamic_constraints(
        model=model, x=x, s={},
        config=config,
        constraint_df=constraint_df,
        namelist_df=namelist_df,
        partners_df=partners_df,
        last_month_df=last_month_df,
        fix_assignment_df=fix_assignment_df,
        planned_df=None,
        row_start=row_start, row_end=row_end,
        date_start_col=date_start_col, date_end_col=date_end_col,
        col_to_date=col_to_date, iso_map=iso_map,
        holiday_cols=holiday_cols, holiday_days=holiday_days,
        fixed_duties=fixed_duties, fixed_standbys=fixed_standbys,
        year=year, month=month, year_old=year_old, month_old=month_old,
        exclusion_keywords=exclusion_keywords,
        is_female_pair=is_female_pair, female_indices=female_indices,
        name_to_row=name_to_row, branch_to_row=branch_to_row,
        is_driver=is_driver, partner_pairs=partner_pairs,
        OFFSET_COL=OFFSET_COL, SCALE=SCALE,
        model_constraints=model_constraints,
        slider_overrides=slider_overrides or {}
    )

    # --------------------------------------------------
    # FAIRNESS OBJECTIVE
    # --------------------------------------------------

    # define point scales and constants
    SBF_BONUS = int(sbf_val * SCALE)

    final_scores = {}
    duty_counts = {}

    sum_new_points = sum(
        int(round(p * SCALE * 2))
        for p in points_df["points"]
    )

    carry_scale = data_bundle.get("carry_scale", 1.0)
    carry_average = data_bundle.get("carry_average", 0.0)

    last_month_scale = carry_scale if carry_scale > 0 else 1.0

    # adjust for any changes to the sheet last month

    manual_adjustments = {}
    if isinstance(last_month_df, pd.DataFrame):
        # iterate from row 115 onwards
        for i in range(114, row_end+ 1):
            raw_name = last_month_df.iat[i, 48]
            if pd.isna(raw_name) or str(raw_name).strip() == "":
                continue
            name = str(raw_name).strip().upper()
            change_text = str(last_month_df.iat[i, 49]).strip().upper() # column AX
            base_val = last_month_df.iat[i, 50] # column AY value

            if not name or name == "NAN": continue

            adj = 0
            if "ADD 1X WD" in change_text:    adj = (1 / last_month_scale)
            elif "ADD 1X F" in change_text:   adj = (1 / last_month_scale)
            elif "ADD 1X WE" in change_text:  adj = (2 / last_month_scale)
            elif "ADD 1X H" in change_text: adj = (2 / last_month_scale)
            elif "MINUS 1X WD" in change_text:    adj = -(1 / last_month_scale)
            elif "MINUS 1X F" in change_text:   adj = -(1 / last_month_scale)
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
                bonus_points += int(round(carry_average * SCALE))

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

            if (r, c) in fixed_duties:
                planned_df.iat[r, c] = "D"
            elif solver.Value(var) == 1:
                planned_df.iat[r, c] = "D"

    # Fill excluded rows (SBF/excused/etc) with "0" on every non-D cell so the
    # output sheet visually shows they are fully blocked for the month.
    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        if any(keyword in status_val for keyword in exclusion_keywords):
            for c in range(date_start_col, date_end_col + 1):
                if planned_df.iat[r, c] != "D":
                    planned_df.iat[r, c] = "0"

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
        p = 2.0 if (d.day in holiday_days_new or day_name in ['sat', 'sun']) else (1.0 if day_name == 'fri' else 1.0)
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
    fixed_standbys = set()

    # Build s variables.
    # Rules:
    #   - Excluded persons (SBF/excused/partner/etc) are fully skipped — no standby either.
    #   - If the person is already on duty (D) that day in planned_df, skip that day.
    #   - Hard X in constraint sheet → skip (not available).
    #   - Manual S in constraint sheet → fixed standby (NewConstant(1)).
    #   - Everyone else gets a free NewBoolVar.
    for r in range(row_start, row_end + 1):
        status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
        is_excluded_for_month = any(keyword in status_val for keyword in exclusion_keywords)

        # Fully excluded people do no standby at all
        if is_excluded_for_month:
            continue

        for c in range(date_start_col, date_end_col + 1):
            # Already assigned as duty this day — cannot also be standby
            if planned_df.iat[r, c] == "D":
                continue

            cell = str(constraint_df.iat[r, c]).strip().upper() if not pd.isna(constraint_df.iat[r, c]) else ""

            # Hard unavailability — skip entirely (no variable needed)
            if cell == "X":
                continue

            # Manual S → fixed standby
            if cell == "S":
                s[(r, c)] = model_s.NewConstant(1)
                fixed_standbys.add((r, c))
                continue

            # Free standby variable
            s[(r, c)] = model_s.NewBoolVar(f"s_{r}_{c}")

    # --------------------------------------------------
    # DYNAMIC STANDBY CONSTRAINTS
    # --------------------------------------------------
    # s is now fully populated — safe to apply constraints
    _sb_soft, _ = apply_dynamic_constraints(
        model=model_s, x={}, s=s,
        config=config,
        constraint_df=constraint_df,
        namelist_df=namelist_df,
        partners_df=partners_df,
        last_month_df=last_month_df,
        fix_assignment_df=fix_assignment_df,
        planned_df=planned_df,
        row_start=row_start, row_end=row_end,
        date_start_col=date_start_col, date_end_col=date_end_col,
        col_to_date=col_to_date, iso_map=iso_map,
        holiday_cols=holiday_cols, holiday_days=holiday_days,
        fixed_duties=fixed_duties, fixed_standbys=fixed_standbys,
        year=year, month=month, year_old=year_old, month_old=month_old,
        exclusion_keywords=exclusion_keywords,
        is_female_pair=is_female_pair, female_indices=female_indices,
        name_to_row=name_to_row, branch_to_row=branch_to_row,
        is_driver=is_driver, partner_pairs=partner_pairs,
        OFFSET_COL=OFFSET_COL, SCALE=SCALE,
        model_constraints=model_constraints,
        slider_overrides=slider_overrides or {}
    )

    # Count S per person and cap at 5; minimise the maximum across all persons
    s_counts = {}
    for r in range(row_start, row_end + 1):
        personal_s_vars = [s[(r, c)] for c in range(date_start_col, date_end_col + 1) if (r, c) in s]
        if not personal_s_vars:
            continue
        count_var = model_s.NewIntVar(0, 5, f"s_count_{r}")
        model_s.Add(count_var == sum(personal_s_vars))
        s_counts[r] = count_var

    if s_counts:
        max_s = model_s.NewIntVar(0, 5, "max_s")
        for count_var in s_counts.values():
            model_s.Add(max_s >= count_var)
        model_s.Minimize(max_s + sum(_sb_soft))

    solver_s = cp_model.CpSolver()
    solver_s.parameters.max_time_in_seconds = 10
    status_s = solver_s.Solve(model_s)

    if status_s in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("Standby pass successful.")
        for (r, c), var in s.items():
            if (r, c) in fixed_standbys or solver_s.Value(var) == 1:
                planned_df.iat[r, c] = "S"
    else:
        print("Could not find a feasible solution for Standby.")

    return planned_df, normal_scale, status, {
        'row_start': row_start,
        'row_end': row_end,
        'date_start_col': date_start_col,
        'date_end_col': date_end_col,
        'constraints_col': constraints_col
    }