import json
from collections import defaultdict
from datetime import datetime

DAY_TYPE_WEEKDAY_NUMS = [0, 1, 2, 3]   # Mon-Thu
DAY_TYPE_FRIDAY_NUM   = 4
DAY_TYPE_WEEKEND_NUMS = [5, 6]

def _day_type_matches(dt_obj, day_type_str, holiday_days):
    d = day_type_str.lower()
    wd = dt_obj.weekday()
    if d == "any":
        return True
    if d == "weekday":
        return wd in DAY_TYPE_WEEKDAY_NUMS
    if d == "friday":
        return wd == DAY_TYPE_FRIDAY_NUM
    if d == "weekend":
        return wd in DAY_TYPE_WEEKEND_NUMS
    if d == "holiday":
        return dt_obj.day in holiday_days
    return False

def _get_last_month_workers_by_day_type(last_month_df, year_old, month_old, 
                                         row_start, row_end, date_start_col, holiday_days_old):
    from datetime import datetime
    import calendar as cal_mod
    workers = defaultdict(set)
    if last_month_df is None:
        return workers
    import pandas as pd
    if not isinstance(last_month_df, pd.DataFrame):
        return workers
    _, last_num_days = cal_mod.monthrange(year_old, month_old)
    for c in range(date_start_col, date_start_col + last_num_days):
        day_num = c - date_start_col + 1
        last_date = datetime(year_old, month_old, day_num)
        for r in range(row_start, row_end + 1):
            try:
                cell_val = str(last_month_df.iat[r, c]).strip().upper()
            except:
                continue
            if cell_val == "D":
                name = str(last_month_df.iat[r, 1]).strip().upper()
                wd = last_date.weekday()
                if wd in DAY_TYPE_WEEKEND_NUMS:
                    workers["weekend"].add(name)
                elif wd == DAY_TYPE_FRIDAY_NUM:
                    workers["friday"].add(name)
                elif day_num in holiday_days_old:
                    workers["holiday"].add(name)
                else:
                    workers["weekday"].add(name)
    return workers

def apply_dynamic_constraints(
    model, x, s, config,
    constraint_df, namelist_df, partners_df, last_month_df,
    fix_assignment_df, planned_df,
    row_start, row_end, date_start_col, date_end_col,
    col_to_date, iso_map, holiday_cols, holiday_days,
    fixed_duties, fixed_standbys,
    year, month, year_old, month_old,
    exclusion_keywords, is_female_pair, female_indices,
    name_to_row, branch_to_row, is_driver, partner_pairs,
    OFFSET_COL, SCALE,
    model_constraints
):
    import pandas as pd
    soft_penalties = []
    has_at_least_one_duty = {}

    # build attribute lookups
    row_to_name = {r: str(constraint_df.iat[r, 1]).strip().upper() for r in range(row_start, row_end + 1)}
    name_to_branch = {}
    name_to_driving = {}
    name_to_partner = {}
    for i in range(len(namelist_df)):
        n = str(namelist_df.iloc[i, 1]).strip().upper()
        b = str(namelist_df.iloc[i, 2]).strip().upper() if len(namelist_df.columns) > 2 else ""
        d = str(namelist_df.iloc[i, 3]).strip().upper() if len(namelist_df.columns) > 3 else ""
        name_to_branch[n] = b
        name_to_driving[n] = d
    for r1, r2 in partner_pairs:
        name_to_partner[row_to_name.get(r1,"")] = row_to_name.get(r2,"")
        name_to_partner[row_to_name.get(r2,"")] = row_to_name.get(r1,"")

    # last month day type workers
    holiday_days_old = set()  # simplified — could be extended
    last_month_workers = _get_last_month_workers_by_day_type(
        last_month_df, year_old, month_old, row_start, row_end, date_start_col, holiday_days_old
    )

    # scalefactor and sbf from model_constraints
    scalefactor = model_constraints.get('scalefactor', 4)
    sbf_val = model_constraints.get('sbf_val', 2)
    hard4_default = model_constraints.get('hard4', 4)
    hard5_default = model_constraints.get('hard5', 3)
    hard1_default = model_constraints.get('hard1', 2)
    hard1s_default = model_constraints.get('hard1s', 2)
    hard2s_default = model_constraints.get('hard2s', 2)

    for cid, cv in config.items():
        if cid.startswith("_"):
            continue
        if not cv.get("active", True):
            continue

        # parse JSON rule from param field
        param_str = cv.get("param", "")
        if not param_str or not param_str.strip().startswith("{"):
            continue
        try:
            rule = json.loads(param_str)
        except:
            continue

        subject    = rule.get("subject", "person")
        assignment = rule.get("assignment", "D").upper()
        cond_type  = rule.get("condition_type", "none")
        cond_val   = rule.get("condition_value", "")
        action     = rule.get("action", "")
        action_val = rule.get("action_value", "")
        is_soft    = rule.get("soft", False)
        penalty    = int(rule.get("penalty", 0))

        # resolve numeric action value
        try:
            action_num = int(action_val) if action_val != "" else 0
        except:
            action_num = 0

        # ── build row filter based on condition ──
        def rows_matching_condition():
            rows = list(range(row_start, row_end + 1))
            if cond_type == "none":
                return rows
            if cond_type == "worked_day_type_last_month":
                workers_set = last_month_workers.get(cond_val.lower(), set())
                return [r for r in rows if row_to_name.get(r,"") in workers_set]
            if cond_type == "person_attribute_is":
                attr = cond_val.lower()
                if attr == "gender=female":
                    return [r for r in rows if is_female_pair.get(r, False)]
                if attr == "partner":
                    return [r for r in rows if row_to_name.get(r,"") in name_to_partner]
                if attr == "driving_status":
                    return rows  # all rows — filtered per action
            return rows

        filtered_rows = rows_matching_condition()

        # ── apply action ──

        if action == "exactly_x_per_day":
            n = action_num if action_num > 0 else (hard1_default if assignment == "D" else hard1s_default)
            if assignment == "D":
                for c in range(date_start_col, date_end_col + 1):
                    model.Add(sum(x[(r, c)] for r in range(row_start, row_end + 1)) == n)
            elif assignment == "S" and s:
                for c in range(date_start_col, date_end_col + 1):
                    day_vars = [s[(r, c)] for r in range(row_start, row_end + 1) if (r, c) in s]
                    if day_vars:
                        model.Add(sum(day_vars) == n)

        elif action == "at_most_x_per_month":
            n = action_num if action_num > 0 else hard5_default
            for r in filtered_rows:
                status_val = str(fix_assignment_df.iat[r, OFFSET_COL + 1]).strip().upper()
                is_excluded = any(k in status_val for k in exclusion_keywords)
                total_vars = [x[(r, c)] for c in range(date_start_col, date_end_col + 1)]
                if is_excluded:
                    num_hol = sum(1 for c in holiday_cols if (r, c) in fixed_duties)
                    model.Add(sum(total_vars) == num_hol)
                elif is_female_pair.get(r, False):
                    model.Add(sum(total_vars) <= n - 1)
                else:
                    model.Add(sum(total_vars) <= n)

        elif action == "at_most_x_per_week":
            n = action_num if action_num > 0 else 1
            for r in filtered_rows:
                for week, cols in iso_map.items():
                    vars_in_week = [x[(r, c)] for c in cols if (r, c) not in fixed_duties or c not in holiday_cols]
                    if vars_in_week:
                        model.Add(sum(vars_in_week) <= n)

        elif action == "gap_at_least_x_days":
            gap = action_num if action_num > 0 else hard4_default
            if assignment == "D":
                for r in filtered_rows:
                    name = row_to_name.get(r, "")
                    from collections import defaultdict as dd
                    duties_lm = dd(list)
                    if last_month_df is not None and isinstance(last_month_df, pd.DataFrame):
                        import calendar as cal_mod
                        _, lnd = cal_mod.monthrange(year_old, month_old)
                        for c in range(date_start_col, date_start_col + lnd):
                            try:
                                if str(last_month_df.iat[r, c]).strip().upper() == "D":
                                    duties_lm[name].append(c - date_start_col + 1)
                            except: pass
                    lm_list = duties_lm.get(name, [])
                    if lm_list:
                        last_d_num = max(lm_list)
                        last_d_date = datetime(year_old, month_old, last_d_num)
                        for c in range(date_start_col, date_end_col + 1):
                            if (r, c) in fixed_duties: continue
                            if (col_to_date[c] - last_d_date).days < gap:
                                model.Add(x[(r, c)] == 0)
                            else:
                                break
                    for c1 in range(date_start_col, date_end_col + 1):
                        d1 = col_to_date[c1]
                        for c2 in range(c1 + 1, date_end_col + 1):
                            if (r, c1) in fixed_duties or (r, c2) in fixed_duties: continue
                            if (col_to_date[c2] - d1).days >= gap: break
                            model.Add(x[(r, c1)] + x[(r, c2)] <= 1)
            elif assignment == "DS" and s:
                gap = action_num if action_num > 0 else hard2s_default
                for r in filtered_rows:
                    for c1 in range(date_start_col, date_end_col + 1):
                        d1 = col_to_date[c1]
                        for c2 in range(c1 + 1, date_end_col + 1):
                            if (col_to_date[c2] - d1).days >= gap: break
                            if planned_df is not None:
                                try:
                                    if planned_df.iat[r, c1] == "D" and (r, c2) in s:
                                        model.Add(s[(r, c2)] == 0)  # type: ignore[index]
                                    if (r, c1) in s and planned_df.iat[r, c2] == "D":
                                        model.Add(s[(r, c1)] == 0)  # type: ignore[index]
                                    if (r, c1) in s and (r, c2) in s:
                                        model.Add(s[(r, c1)] + s[(r, c2)] <= gap - 1)  # type: ignore[index]
                                except: pass

        elif action == "cannot_work_day_type":
            for r in filtered_rows:
                for c in range(date_start_col, date_end_col + 1):
                    if (r, c) in fixed_duties: continue
                    dt_obj = col_to_date[c]
                    if _day_type_matches(dt_obj, action_val, holiday_days):
                        if (r, c) in x:
                            model.Add(x[(r, c)] == 0)

        elif action == "grouped_0_or_n":
            n = action_num if action_num > 0 else 2
            if cond_type == "person_attribute_is" and "gender=female" in cond_val:
                for c in range(date_start_col, date_end_col + 1):
                    fvars = [x[(r, c)] for r in female_indices if (r, c) in x]
                    if fvars:
                        fc = model.NewIntVar(0, n, f"grouped_{cid}_{c}")
                        model.Add(fc == sum(fvars))
                        is_grp = model.NewBoolVar(f"is_grp_{cid}_{c}")
                        model.Add(fc == n).OnlyEnforceIf(is_grp)
                        model.Add(fc == 0).OnlyEnforceIf(is_grp.Not())

        elif action == "match_branch_of_d" and s:
            from collections import Counter
            for c in range(date_start_col, date_end_col + 1):
                if planned_df is None: continue
                try:
                    d_rows_day = [r for r in range(row_start, row_end + 1) if planned_df.iat[r, c] == "D"]
                except: continue
                num_d = len(d_rows_day)
                day_s_vars = [s[(r, c)] for r in range(row_start, row_end + 1) if (r, c) in s]
                if num_d == 0:
                    for v in day_s_vars: model.Add(v == 0)
                    continue
                req_branches = []
                for ri in d_rows_day:
                    nm = row_to_name.get(ri, "")
                    br = name_to_branch.get(nm, "")
                    if br: req_branches.append(br)
                bc = Counter(req_branches)
                if len(day_s_vars) >= num_d:
                    model.Add(sum(day_s_vars) == num_d)
                for branch, rs in bc.items():
                    bvars = [s[(r, c)] for r in branch_to_row.get(branch, []) if (r, c) in s]
                    if len(bvars) >= rs:
                        model.Add(sum(bvars) == rs)
                for branch, rows_b in branch_to_row.items():
                    if branch not in bc:
                        for r in rows_b:
                            if (r, c) in s: model.Add(s[(r, c)] == 0)

        elif action == "prefer_together":
            if action_val == "partner":
                for r1, r2 in partner_pairs:
                    for c in range(date_start_col, date_end_col + 1):
                        if (r1, c) in x and (r2, c) in x:
                            is_split = model.NewBoolVar(f"split_{cid}_{r1}_{r2}_{c}")
                            model.Add(x[(r1, c)] - x[(r2, c)] <= is_split)
                            model.Add(x[(r2, c)] - x[(r1, c)] <= is_split)
                            soft_penalties.append(is_split * penalty)

        elif action == "prefer_different":
            if action_val == "branch":
                for c in range(date_start_col, date_end_col + 1):
                    for branch, rows_b in branch_to_row.items():
                        bvars = [x[(r, c)] for r in rows_b if (r, c) in x]
                        if len(bvars) >= 2:
                            viol = model.NewBoolVar(f"branch_viol_{cid}_{branch}_{c}")
                            model.Add(sum(bvars) < 2).OnlyEnforceIf(viol.Not())
                            model.Add(sum(bvars) == 2).OnlyEnforceIf(viol)
                            soft_penalties.append(viol * penalty)

        elif action == "prefer_balanced_mix":
            if action_val == "driving_status":
                for c in range(date_start_col, date_end_col + 1):
                    drivers = [x[(r, c)] for r in range(row_start, row_end + 1) if (r, c) in x and r in is_driver]
                    if drivers:
                        dc = model.NewIntVar(0, 2, f"dcount_{cid}_{c}")
                        model.Add(dc == sum(drivers))
                        mm = model.NewBoolVar(f"dmm_{cid}_{c}")
                        model.Add(dc != 1).OnlyEnforceIf(mm)
                        model.Add(dc == 1).OnlyEnforceIf(mm.Not())
                        soft_penalties.append(mm * penalty)

        elif action == "at_least_1_per_month":
            for r in filtered_rows:
                has_at_least_one_duty[r] = model.NewBoolVar(f"has_duty_{cid}_{r}")
                ds = sum(x[(r, c)] for c in range(date_start_col, date_end_col + 1) if (r, c) in x)
                model.Add(ds >= 1).OnlyEnforceIf(has_at_least_one_duty[r])
                model.Add(ds == 0).OnlyEnforceIf(has_at_least_one_duty[r].Not())
                soft_penalties.append(has_at_least_one_duty[r].Not() * penalty)

    return soft_penalties, has_at_least_one_duty