import json
from collections import defaultdict
from datetime import datetime
import calendar as cal_mod

DAY_TYPE_MAP = {
    "weekday": [0,1,2,3],
    "friday":  [4],
    "weekend": [5,6],
}

def _matches_day_type(dt_obj, day_type, holiday_days):
    d = day_type.lower()
    wd = dt_obj.weekday()
    if d == "any":     return True
    if d == "holiday": return dt_obj.day in holiday_days
    return wd in DAY_TYPE_MAP.get(d, [])

def _last_month_worked(last_month_df, year_old, month_old,
                       row_start, row_end, date_start_col, holiday_days_old):
    """Returns dict: day_type_str -> set of names who worked that type last month."""
    import pandas as pd
    workers = defaultdict(set)
    if last_month_df is None or not isinstance(last_month_df, pd.DataFrame):
        return workers
    _, lnd = cal_mod.monthrange(year_old, month_old)
    for c in range(date_start_col, date_start_col + lnd):
        day_num = c - date_start_col + 1
        ld = datetime(year_old, month_old, day_num)
        for r in range(row_start, row_end + 1):
            try:
                cv = str(last_month_df.iat[r, c]).strip().upper()
            except:
                continue
            if cv == "D":
                name = str(last_month_df.iat[r, 1]).strip().upper()
                wd = ld.weekday()
                if day_num in holiday_days_old:
                    workers["holiday"].add(name)
                elif wd in [5,6]:
                    workers["weekend"].add(name)
                elif wd == 4:
                    workers["friday"].add(name)
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
    model_constraints,
    slider_overrides=None
):
    import pandas as pd
    soft_penalties = []
    has_at_least_one_duty = {}

    # ── build attribute lookups ──
    row_to_name = {r: str(constraint_df.iat[r,1]).strip().upper()
                   for r in range(row_start, row_end+1)}
    name_to_branch  = {}
    name_to_driving = {}
    for i in range(len(namelist_df)):
        n = str(namelist_df.iloc[i,1]).strip().upper()
        b = str(namelist_df.iloc[i,2]).strip().upper() if len(namelist_df.columns)>2 else ""
        d = str(namelist_df.iloc[i,3]).strip().upper() if len(namelist_df.columns)>3 else ""
        name_to_branch[n]  = b
        name_to_driving[n] = d

    # partner row pairs already built outside
    partner_row_set = set()
    for r1,r2 in partner_pairs:
        partner_row_set.add((min(r1,r2), max(r1,r2)))

    # last month worked by day type (used by CLASS: ALLOW)
    lm_workers = _last_month_worked(
        last_month_df, year_old, month_old,
        row_start, row_end, date_start_col, set()
    )

    # model_constraints fallbacks
    hard4  = model_constraints.get('hard4', 4)

    if slider_overrides is None:
        slider_overrides = {}

    for cid, cv in config.items():
        if cid.startswith("_"):
            continue
        if not cv.get("active", True):
            continue
        param_str = cv.get("param","")
        if not param_str or not param_str.strip().startswith("{"):
            continue
        try:
            rule = json.loads(param_str)
        except:
            continue

        # apply slider override — override numeric param at runtime
        if cid in slider_overrides:
            override_val = slider_overrides[cid]
            cls_check = rule.get("class","")
            if cls_check == "value":
                rule["number"] = override_val
            elif cls_check == "gap":
                rule["days"] = override_val
            elif cls_check in ("grouping","allow"):
                rule["penalty"] = override_val

        cls     = rule.get("class","")
        is_soft = rule.get("soft", False)
        penalty = int(rule.get("penalty", 0))

        # ════════════════════════════════
        # CLASS: VALUE
        # ════════════════════════════════
        if cls == "value":
            subj1    = rule.get("subject1","person")
            operator = rule.get("operator","=")
            number   = int(rule.get("number", 1))
            subj2    = rule.get("subject2","D").upper()
            per      = rule.get("per","month")

            if subj2 == "D" and not x:
                continue

            if subj1 == "day":
                # each day must have op N of subj2
                if subj2 == "D":
                    for c in range(date_start_col, date_end_col+1):
                        day_vars = [x[(r,c)] for r in range(row_start, row_end+1)]
                        if operator == "=":
                            model.Add(sum(day_vars) == number)
                        elif operator == "<=":
                            model.Add(sum(day_vars) <= number)
                        elif operator == ">=":
                            if is_soft:
                                viol = model.NewBoolVar(f"val_viol_{cid}_{c}")
                                model.Add(sum(day_vars) >= number).OnlyEnforceIf(viol.Not())
                                soft_penalties.append(viol * penalty)
                            else:
                                model.Add(sum(day_vars) >= number)
                elif subj2 == "S" and s:
                    for c in range(date_start_col, date_end_col+1):
                        day_vars = [s[(r,c)] for r in range(row_start, row_end+1) if (r,c) in s]
                        if day_vars:
                            if operator == "=":
                                model.Add(sum(day_vars) == number)
                            elif operator == "<=":
                                model.Add(sum(day_vars) <= number)
                            elif operator == ">=":
                                model.Add(sum(day_vars) >= number)

            elif subj1 == "person":
                if per == "week" and subj2 == "D":
                    for r in range(row_start, row_end+1):
                        for week, cols in iso_map.items():
                            wvars = [x[(r,c)] for c in cols
                                if (r,c) not in fixed_duties]
                            if wvars:
                                if operator == "<=":
                                    model.Add(sum(wvars) <= number)
                                elif operator == "=":
                                    model.Add(sum(wvars) == number)

                elif per == "month" and subj2 == "D":
                    for r in range(row_start, row_end+1):
                        total = [x[(r,c)] for c in range(date_start_col, date_end_col+1)
                            if (r,c) not in fixed_duties]
                        if operator == "<=" and not is_soft:
                                model.Add(sum(total) <= number)
                        elif operator == ">=" and is_soft:
                            has_at_least_one_duty[r] = model.NewBoolVar(f"has_duty_{cid}_{r}")
                            ds = sum(x[(r,c)] for c in range(date_start_col, date_end_col+1) if (r,c) in x)
                            model.Add(ds >= number).OnlyEnforceIf(has_at_least_one_duty[r])
                            model.Add(ds == 0).OnlyEnforceIf(has_at_least_one_duty[r].Not())
                            soft_penalties.append(has_at_least_one_duty[r].Not() * penalty)

        # ════════════════════════════════
        # CLASS: ALLOW
        # ════════════════════════════════
        elif cls == "allow":
            cond_dt   = rule.get("condition_day_type","weekend")
            logic     = rule.get("logic","cannot")
            action_dt = rule.get("action_day_type","weekend")
            workers   = lm_workers.get(cond_dt.lower(), set())

            if not x:
                continue

            for r in range(row_start, row_end+1):
                name = row_to_name.get(r,"")
                if name not in workers:
                    continue
                for c in range(date_start_col, date_end_col+1):
                    if (r,c) in fixed_duties:
                        continue
                    dt_obj = col_to_date[c]
                    if _matches_day_type(dt_obj, action_dt, holiday_days):
                        if logic == "cannot" and (r,c) in x:
                            model.Add(x[(r,c)] == 0)
                        elif logic == "can":
                            pass  # no restriction

        # ════════════════════════════════
        # CLASS: GAP
        # ════════════════════════════════
        elif cls == "gap":
            from_type = rule.get("from_type","D").upper()
            to_type   = rule.get("to_type","D").upper()
            days      = int(rule.get("days", hard4))

            if from_type == "D" and to_type == "D" and not x:
                continue

            if from_type == "D" and to_type == "D":
                # cross-month and internal D-D gap
                for r in range(row_start, row_end+1):
                    # cross-month
                    lm_duties = []
                    if last_month_df is not None and isinstance(last_month_df, pd.DataFrame):
                        _, lnd = cal_mod.monthrange(year_old, month_old)
                        for c in range(date_start_col, date_start_col+lnd):
                            try:
                                if str(last_month_df.iat[r,c]).strip().upper() == "D":
                                    lm_duties.append(c - date_start_col + 1)
                            except: pass
                    if lm_duties:
                        last_d = datetime(year_old, month_old, max(lm_duties))
                        for c in range(date_start_col, date_end_col+1):
                            if (r,c) in fixed_duties:
                                continue
                            if (col_to_date[c] - last_d).days < days:
                                model.Add(x[(r,c)] == 0)
                            else:
                                break
                    # internal
                    for c1 in range(date_start_col, date_end_col+1):
                        d1 = col_to_date[c1]
                        for c2 in range(c1+1, date_end_col+1):
                            if (r,c1) in fixed_duties or (r,c2) in fixed_duties: continue
                            if (col_to_date[c2]-d1).days >= days: break
                            model.Add(x[(r,c1)] + x[(r,c2)] <= 1)

            elif from_type in ("D","S") and to_type in ("D","S") and s:
                # D-S, S-D, or S-S gaps
                for r in range(row_start, row_end+1):
                    for c1 in range(date_start_col, date_end_col+1):
                        d1 = col_to_date[c1]
                        for c2 in range(c1+1, date_end_col+1):
                            if (col_to_date[c2]-d1).days >= days: break
                            try:
                                c1_is_d = planned_df is not None and planned_df.iat[r,c1]=="D"
                                c2_is_d = planned_df is not None and planned_df.iat[r,c2]=="D"
                            except: continue
                            c1_is_s = (r,c1) in s
                            c2_is_s = (r,c2) in s

                            if from_type=="D" and to_type=="S":
                                if c1_is_d and c2_is_s:
                                    model.Add(s[(r,c2)] == 0)
                                if c1_is_s and c2_is_d:
                                    model.Add(s[(r,c1)] == 0)
                            elif from_type=="S" and to_type=="S":
                                if c1_is_s and c2_is_s:
                                    model.Add(s[(r,c1)] + s[(r,c2)] <= 1)

        # ════════════════════════════════
        # CLASS: GROUPING
        # ════════════════════════════════
        elif cls == "grouping":
            trait  = rule.get("trait","")
            logic  = rule.get("logic","must")

            if not s and rule.get("duty_type","D").upper() == "S":
                continue

            if trait == "same_gender" and logic == "must":
                # females must be together: count is 0 or len(female_indices)
                n = len(female_indices) if female_indices else 2
                for c in range(date_start_col, date_end_col+1):
                    fvars = [x[(r,c)] for r in female_indices if (r,c) in x]
                    if fvars:
                        fc = model.NewIntVar(0, n, f"fcount_{cid}_{c}")
                        model.Add(fc == sum(fvars))
                        is_grp = model.NewBoolVar(f"fgrp_{cid}_{c}")
                        model.Add(fc == n).OnlyEnforceIf(is_grp)
                        model.Add(fc == 0).OnlyEnforceIf(is_grp.Not())

            elif trait == "same_branch" and logic == "must_match_d" and s:
                # S must match D branch on same day
                from collections import Counter
                for c in range(date_start_col, date_end_col+1):
                    if planned_df is None: continue
                    try:
                        d_rows_day = [r for r in range(row_start, row_end+1)
                                      if planned_df.iat[r,c]=="D"]
                    except: continue
                    num_d = len(d_rows_day)
                    day_s_vars = [s[(r,c)] for r in range(row_start, row_end+1) if (r,c) in s]
                    if num_d == 0:
                        for v in day_s_vars: model.Add(v==0)
                        continue
                    req_br = []
                    for ri in d_rows_day:
                        nm = row_to_name.get(ri,"")
                        br = name_to_branch.get(nm,"")
                        if br: req_br.append(br)
                    bc = Counter(req_br)
                    if len(day_s_vars) >= num_d:
                        model.Add(sum(day_s_vars) == num_d)
                    for branch, rs in bc.items():
                        bvars = [s[(r,c)] for r in branch_to_row.get(branch,[]) if (r,c) in s]
                        if len(bvars) >= rs:
                            model.Add(sum(bvars) == rs)
                    for branch, rows_b in branch_to_row.items():
                        if branch not in bc:
                            for r in rows_b:
                                if (r,c) in s: model.Add(s[(r,c)]==0)

            elif trait == "partners" and logic == "must":
                for r1,r2 in partner_pairs:
                    for c in range(date_start_col, date_end_col+1):
                        if (r1,c) in x and (r2,c) in x:
                            if (r1,c) in fixed_duties or (r2,c) in fixed_duties:
                                continue
                            is_split = model.NewBoolVar(f"split_{cid}_{r1}_{r2}_{c}")
                            model.Add(x[(r1,c)]-x[(r2,c)] <= is_split)
                            model.Add(x[(r2,c)]-x[(r1,c)] <= is_split)
                            soft_penalties.append(is_split * penalty)

            elif trait == "same_branch" and logic == "cannot":
                for c in range(date_start_col, date_end_col+1):
                    for branch, rows_b in branch_to_row.items():
                        bvars = [x[(r,c)] for r in rows_b
                            if (r,c) in x and (r,c) not in fixed_duties]
                        if len(bvars) >= 2:
                            viol = model.NewBoolVar(f"br_viol_{cid}_{branch}_{c}")
                            model.Add(sum(bvars) < 2).OnlyEnforceIf(viol.Not())
                            model.Add(sum(bvars) == 2).OnlyEnforceIf(viol)
                            soft_penalties.append(viol * penalty)

            elif trait == "drivers" and logic == "cannot":
                for c in range(date_start_col, date_end_col+1):
                    drivers = [x[(r,c)] for r in range(row_start, row_end+1)
                        if (r,c) in x and r in is_driver and (r,c) not in fixed_duties]
                    if drivers:
                        dc = model.NewIntVar(0, len(drivers), f"dcount_{cid}_{c}")
                        model.Add(dc == sum(drivers))
                        mm = model.NewBoolVar(f"dmm_{cid}_{c}")
                        model.Add(dc != 1).OnlyEnforceIf(mm)
                        model.Add(dc == 1).OnlyEnforceIf(mm.Not())
                        soft_penalties.append(mm * penalty)

    return soft_penalties, has_at_least_one_duty