import streamlit as st
import pandas as pd
import gspread
import planner_engine
import user_engine
import traceback
import calendar
from datetime import date, timedelta
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from ortools.sat.python import cp_model
import requests as http_requests

# Passwords now stored in CONFIG sheet

st.set_page_config(page_title="Duty Planner", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# --------------------------------------------------
# AUTHENTICATION
# --------------------------------------------------

def get_gspread_auth():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = service_account.Credentials.from_service_account_info(
                creds_dict,
                scopes=SCOPES
            )
            client = gspread.authorize(creds)
            return client
        else:
            st.error("❌ 'gcp_service_account' not found in secrets.toml")
            st.stop()
    except Exception as e:
        st.error(f"❌ Authentication Error: {e}")
        st.stop()

def get_personal_drive_service():
    info = st.secrets["personal_account"]
    creds = Credentials(
        token=info["token"],
        refresh_token=info["refresh_token"],
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        token_uri=info["token_uri"]
    )
    if creds.expired:
        creds.refresh(Request())
    return build('drive', 'v3', credentials=creds)

# --------------------------------------------------
# CACHED DATA FETCHERS
# --------------------------------------------------

@st.cache_data(ttl=300, show_spinner=False)
def fetch_sheet_data(_client, spreadsheet_name, sheet_name):
    sh = _client.open(spreadsheet_name)
    return sh.worksheet(sheet_name).get_all_values()

@st.cache_data(ttl=300, show_spinner=False)
def fetch_namelist(_client, spreadsheet_name):
    try:
        sh = _client.open(spreadsheet_name)
        ws = sh.worksheet("Namelist")
        records = ws.get_all_records()
        return [r['NAME'] for r in records if r.get('NAME')]
    except Exception as e:
        print(f"Error fetching namelist: {e}")
        return []

@st.cache_data(ttl=300, show_spinner=False)
def fetch_trait_options(_client, spreadsheet_name):
    """Return the ordered list of trait labels managed by Dev (stored in CONFIG _TRAITS row)."""
    try:
        sh = _client.open(spreadsheet_name)
        ws = sh.worksheet("CONFIG")
        for row in ws.get_all_values():
            if row and row[0].strip() == "_TRAITS":
                raw = row[1].strip() if len(row) > 1 else ""
                return [t.strip() for t in raw.split(",") if t.strip()]
        return []
    except Exception as e:
        print(f"Error fetching trait options: {e}")
        return []

@st.cache_data(ttl=600, show_spinner=False)
def fetch_spreadsheet_id(_personal_drive, folder_id, spreadsheet_name):
    gs_query = (
        f"name = '{spreadsheet_name}' "
        f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
        f"and trashed = false "
        f"and '{folder_id}' in parents"
    )
    results = _personal_drive.files().list(q=gs_query, fields="files(id)").execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

# --------------------------------------------------
# CONVERT FILES FROM .XLSX TO SHEETS
# --------------------------------------------------


@st.cache_data(ttl=60, show_spinner=False)
def fetch_config(_client, spreadsheet_name):
    try:
        sh = _client.open(spreadsheet_name)
        ws = sh.worksheet("CONFIG")
        rows = ws.get_all_values()
        cfg = {}
        pwd = {}
        in_passwords = False
        for row in rows:
            if not any(row):
                continue
            if row[0].strip().upper() == "KEY":
                in_passwords = True
                continue
            if in_passwords:
                key = row[0].strip()
                val = row[1].strip() if len(row) > 1 else ""
                if key:
                    pwd[key] = val
            else:
                cid = row[0].strip()
                if not cid or cid.upper() == "CONSTRAINT_ID":
                    continue
                import json as _json
                _raw_param = row[5].strip() if len(row) > 5 else ""
                try:
                    _rule = _json.loads(_raw_param) if _raw_param.startswith("{") else {}
                except:
                    _rule = {}
                cfg[cid] = {
                    "label":        row[1].strip() if len(row) > 1 else cid,
                    "type":         row[2].strip() if len(row) > 2 else "",
                    "active":       row[3].strip().upper() == "TRUE" if len(row) > 3 else True,
                    "draft_active": row[4].strip().upper() == "TRUE" if len(row) > 4 else True,
                    "param":        _raw_param,
                    "rule":         _rule,
                    "param_label":  row[6].strip() if len(row) > 6 else "",
                    "duty_type":    row[7].strip() if len(row) > 7 else "",
                    "class":        row[8].strip() if len(row) > 8 else "",
                    "description":  row[9].strip() if len(row) > 9 else "",
                }
        cfg["_passwords"] = pwd
        return cfg
    except Exception as e:
        return {"_passwords": {"admin_password": "password", "user_password": "weapons"}, "_error": str(e)}

def convert_if_excel(client, spreadsheet_name):
    personal_drive = get_personal_drive_service()
    folder_id = st.secrets["app_config"]["personal_drive_folder_id"]

    # check if Google Sheet already exists in personal folder
    gs_query = (
        f"name = '{spreadsheet_name}' "
        f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
        f"and trashed = false "
        f"and '{folder_id}' in parents"
    )
    gs_results = personal_drive.files().list(q=gs_query, fields="files(id)").execute()
    gs_files = gs_results.get('files', [])

    if gs_files:
        return client.open_by_key(gs_files[0]['id'])

    # check if Excel file exists in personal folder
    ex_query = (
        f"(name = '{spreadsheet_name}' or name = '{spreadsheet_name}.xlsx') "
        f"and mimeType != 'application/vnd.google-apps.spreadsheet' "
        f"and trashed = false "
        f"and '{folder_id}' in parents"
    )
    ex_results = personal_drive.files().list(q=ex_query, fields="files(id, name)").execute()
    ex_files = ex_results.get('files', [])

    if not ex_files:
        raise FileNotFoundError(f"Could not find any active Excel or Google Sheet named '{spreadsheet_name}' in your Drive folder")

    with st.spinner("📦 Excel source detected. Converting to Google Sheets..."):
        converted_file = personal_drive.files().copy(
            fileId=ex_files[0]['id'],
            body={
                'name': spreadsheet_name,
                'mimeType': 'application/vnd.google-apps.spreadsheet',
                'parents': [folder_id]
            }
        ).execute()

    return client.open_by_key(converted_file['id'])

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_role' not in st.session_state:
    st.session_state['user_role'] = None

# --------------------------------------------------
# LOGIN
# --------------------------------------------------

def logout():
    st.session_state['logged_in'] = False
    st.session_state['user_role'] = 'User'

if not st.session_state['logged_in']:
    st.title("🚀 Duty Planner")
    st.write("Please select your access level to continue.")
    try:
        _login_client = get_gspread_auth()
        _login_cfg = fetch_config(_login_client, "MASTER SHEET")
        _admin_pw = _login_cfg.get("_passwords", {}).get("admin_password", "password")
        _user_pw  = _login_cfg.get("_passwords", {}).get("user_password", "weapons")
    except:
        _admin_pw = "password"
        _user_pw  = "weapons"

    st.subheader("User")
    user_pwd = st.text_input("Enter Password", type="password", key="user_password")
    if st.button("Login as User", use_container_width=True):
        if user_pwd == _user_pw:
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'User'
            st.rerun()
        else:
            st.error("❌ Incorrect Password")

    st.subheader("Admin")
    admin_pwd = st.text_input("Enter Password", type="password", key="admin_password")
    if st.button("Login as Admin", use_container_width=True):
        if admin_pwd == _admin_pw:
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'Admin'
            st.rerun()
        elif admin_pwd == "devpass":
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'Dev'
            st.rerun()
        else:
            st.error("❌ Incorrect Password")
    st.stop()

role = st.session_state['user_role']
if role != 'Dev':
    st.sidebar.button("Logout", on_click=logout, key="admin_logout")

# --------------------------------------------------
# ADMIN INTERFACE
# --------------------------------------------------

if "hard4_initialised" not in st.session_state:
    st.session_state["hard4_slider"] = 4
    st.session_state["hard4_input"] = 4
    st.session_state["hard4_initialised"] = True

if role == 'Admin':

    # --------------------------------------------------
    # ADMIN PAGE NAVIGATION
    # --------------------------------------------------

    st.title("🚀 Duty Planner")

    admin_page = st.sidebar.segmented_control(
        "",
        options=["🗓 Planning", "✏️ Editing"],
    )

    try:
        client = get_gspread_auth()
    except Exception as e:
        st.error(f"❌ Connection Error: {e}")
        st.stop()

    # --------------------------------------------------
    # PLANNING PAGE
    # --------------------------------------------------

    if admin_page == "🗓 Planning":
    
        st.title ("🗓 Planning")

        # planning parameters

        def update_slider(key):
            st.session_state[key + "_slider"] = st.session_state[key + "_input"]

        def update_input(key):
            st.session_state[key + "_input"] = st.session_state[key + "_slider"]

        st.sidebar.title("📅 Planning Settings")

        # ── Load CONFIG sheet ──
        sheet_cfg = fetch_config(client, "MASTER SHEET")
        config = sheet_cfg

        # ── Dynamic constraint sliders from CONFIG ──
        # slider_overrides: cid -> numeric value (overrides CONFIG param at runtime)
        slider_overrides = {}

        def _get_rule_num(cv):
            import json as _j
            try:
                rule = cv.get("rule", {})
                cls  = rule.get("class","")
                soft = rule.get("soft", False)
                if cls == "value":   return int(rule.get("number", 1))
                if cls == "gap":     return int(rule.get("days", 1))
                if cls in ("grouping","allow") and soft:
                    return int(rule.get("penalty", 0))
                # hard grouping and hard allow have no adjustable param
                return None
            except:
                return None

        def _get_rule_num_label(cv):
            """Human label for the numeric param."""
            try:
                rule = cv.get("rule", {})
                cls = rule.get("class","")
                if cls == "value":   return cv.get("label","") + " (number)"
                if cls == "gap":     return cv.get("label","") + " (days)"
                if cls in ("grouping","allow"): return cv.get("label","") + " (penalty)"
            except: pass
            return cv.get("label","")

        hard_constraints = {k: v for k, v in sheet_cfg.items()
                            if not k.startswith("_") and v.get("type","").lower() == "hard"
                            and v.get("active", True)}
        soft_constraints = {k: v for k, v in sheet_cfg.items()
                            if not k.startswith("_") and v.get("type","").lower() == "soft"
                            and v.get("active", True)}

        st.sidebar.subheader("🔒 Hard Constraints")
        for cid, cv in hard_constraints.items():
            default_num = _get_rule_num(cv)
            if default_num is not None:
                sk = f"dyn_slider_{cid}"
                max_val = max(default_num * 3, 20)
                val = st.sidebar.slider(cv.get("label", cid), 0, max_val,
                                        value=st.session_state.get(sk, default_num),
                                        key=sk, step=1)
                slider_overrides[cid] = val
            else:
                st.sidebar.caption(f"✅ {cv.get('label', cid)}")

        st.sidebar.markdown("---")
        st.sidebar.subheader("🔓 Soft Constraints")
        for cid, cv in soft_constraints.items():
            default_num = _get_rule_num(cv)
            if default_num is not None:
                sk = f"dyn_slider_{cid}"
                max_val = max(default_num * 3, 100)
                val = st.sidebar.slider(cv.get("label", cid), 0, max_val,
                                        value=st.session_state.get(sk, default_num),
                                        key=sk, step=10)
                slider_overrides[cid] = val
            else:
                st.sidebar.caption(f"✅ {cv.get('label', cid)}")

        st.sidebar.markdown("---")
        st.sidebar.subheader("💯 Point Allocations")

        if "weekday_slider" not in st.session_state:
            st.session_state["weekday_slider"] = 1.0
        if "friday_slider" not in st.session_state:
            st.session_state["friday_slider"] = 1.0
        if "weekend_slider" not in st.session_state:
            st.session_state["weekend_slider"] = 2.0
        if "holiday_slider" not in st.session_state:
            st.session_state["holiday_slider"] = 2.0

        weekday_val = st.sidebar.slider("Weekday Points", 0.0, 10.0, key="weekday_slider", step=0.5)
        friday_val  = st.sidebar.slider("Friday Points",  0.0, 10.0, key="friday_slider",  step=0.5)
        weekend_val = st.sidebar.slider("Weekend Points", 0.0, 10.0, key="weekend_slider", step=0.5)
        holiday_val = st.sidebar.slider("Holiday Points", 0.0, 10.0, key="holiday_slider", step=0.5)

        point_allocations = {
            "weekday_points": weekday_val,
            "friday_points":  friday_val,
            "weekend_points": weekend_val,
            "holiday_points": holiday_val
        }

        st.sidebar.markdown("---")
        st.sidebar.subheader("⚙️ Optimiser Settings")
        if "scalefactor_slider" not in st.session_state:
            st.session_state["scalefactor_slider"] = 4
        if "sbf_slider" not in st.session_state:
            st.session_state["sbf_slider"] = 2
        scalefactor_val = st.sidebar.slider("Normalisation Scale", 0, 5, key="scalefactor_slider", step=1)
        sbf_val         = st.sidebar.slider("SBF Bonus",           0, 5, key="sbf_slider",         step=1)

        model_constraints = {
            "scalefactor": scalefactor_val,
            "sbf_val":     sbf_val
        }

        # main interface

        _now = date.today(); _opts = [f"{m:02d}{str(y)[2:]}" for y in [_now.year, _now.year+1] for m in range(1,13)]
        mmyy = st.selectbox("Month/Year (MMYY) to plan", options=_opts, key="plan_mmyy")
        spreadsheet_name = f"MASTER SHEET"

        curr_m, curr_y = int(mmyy[:2]), int(mmyy[2:])
        m_new = curr_m + 1 
        if m_new > 12:
            m_new = 1
            y_new = curr_y + 1
        else:
            y_new = curr_y
        m_old = curr_m - 1
        if curr_m == 1:
            m_old = 12
            y_old = curr_y - 1
        else:
            y_old = curr_y

        st.info(f"Planning **{mmyy}**!")

        try:
            personal_drive = get_personal_drive_service()
            folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
            sheet_id = fetch_spreadsheet_id(personal_drive, folder_id, spreadsheet_name)
            if sheet_id:
                st.success(f"✅ Connected to storage!")
            else:
                st.warning(f"⚠️ Connection error: storage failed!")
        except Exception as e:
            st.error(f"❌ Storage failed: {e}")
        
        col1, col2 = st.columns([2,3])

        with col1:
            st.write("Step 1:")

        with col2:
            if st.button("🔥 Run Optimiser"):
                try:

                    sh = convert_if_excel(client, spreadsheet_name)

                    # read average and scale from current month C sheet
                    try:
                        c_ws = sh.worksheet(f"{mmyy}C")
                        avg_val = c_ws.acell('AU2').value
                        scale_val = c_ws.acell('AU3').value
                        carry_average = float(avg_val) if avg_val else 0.0
                        carry_scale = float(scale_val) if scale_val else 1.0
                    except:
                        carry_average = 0.0
                        carry_scale = 1.0

                    with st.spinner("📥 Fetching Sheet Data..."):

                        def get_df(sheet_name, header_row=0, use_cols=None):
                            try:
                                data = sh.worksheet(sheet_name).get_all_values()
                                df = pd.DataFrame(data)
                                df.columns = df.iloc[header_row]
                                df = df[header_row + 1:].reset_index(drop=True)
                                if use_cols:
                                    df = df.iloc[:, :use_cols]
                                return df.head(250)
                            except Exception as e:
                                raise ValueError(f"Error loading sheet '{sheet_name}': {e}")

                        # C sheet: header is row 3 (index 2), offset col is AQ (index 42)
                        constraints_raw = get_df(f"{mmyy}C", header_row=2)
                        constraints_raw.iloc[:, 42] = pd.to_numeric(constraints_raw.iloc[:, 42], errors='coerce').fillna(0)
                        holidays_raw = get_df("Holiday", header_row=0)
                        partners_raw = get_df("Partners", header_row=0, use_cols=3)
                        namelist_raw = get_df("Namelist", header_row=0, use_cols=5)

                        try:
                            last_month_raw = get_df(f"{m_old:02d}{y_old:02d}D", header_row=2)
                        except:
                            st.warning("⚠️ Previous month data not found.")
                            last_month_raw = None
                    

                    with st.spinner("🧠 Solving Optimisation..."):

                        data_bundle = {
                            "constraints": constraints_raw,
                            "holidays": holidays_raw,
                            "year": 2000 + int(mmyy[2:]),
                            "year_old": 2000 + y_old,
                            "month": int(mmyy[:2]),
                            "month_old": m_old,
                            "partners": partners_raw,
                            "namelist": namelist_raw,
                            "last_month": last_month_raw,
                            "carry_average": carry_average,
                            "carry_scale": carry_scale
                        }

                        planned_df, n_scale, status, status_s, ranges = planner_engine.run_optimisation(data_bundle, config, point_allocations, model_constraints, slider_overrides)

                        if planned_df is not None:
                            st.session_state['planned_df'] = planned_df
                            st.session_state['n_scale'] = n_scale
                            st.session_state['ranges'] = ranges
                            st.session_state['active_sh_name'] = sh.title

                            # if the next month falls in a new year, duplicate the
                            # current year's summary sheet and update it for the new year
                            if curr_m == 12:
                                curr_year_full = 2000 + curr_y
                                next_year_full = curr_year_full + 1
                                curr_year_str = str(curr_year_full)
                                next_year_str = str(next_year_full)

                                with st.spinner(f"📅 Creating {next_year_str} sheet..."):
                                    try:
                                        # check if the new year sheet already exists
                                        existing_names = [ws.title for ws in sh.worksheets()]
                                        if next_year_str in existing_names:
                                            st.info(f"ℹ️ Sheet '{next_year_str}' already exists — skipping duplication.")
                                        else:
                                            # find and duplicate the current year sheet
                                            curr_year_ws = sh.worksheet(curr_year_str)
                                            all_sheets = sh.worksheets()
                                            last_index = len(all_sheets)
                                            new_year_ws = sh.duplicate_sheet(
                                                curr_year_ws.id,
                                                insert_sheet_index=last_index,
                                                new_sheet_name=next_year_str
                                            )
                                            # update the year label cell
                                            new_year_ws.update_acell('BM73', '')
                                            new_year_ws.update_acell('BM73', next_year_full)
                                            st.success(f"✅ Created '{next_year_str}' sheet from '{curr_year_str}'!")
                                    except gspread.exceptions.WorksheetNotFound:
                                        st.warning(f"⚠️ Sheet '{curr_year_str}' not found — skipping year sheet creation.")
                                    except Exception as e:
                                        st.warning(f"⚠️ Could not create {next_year_str} sheet: {e}")

                            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                                if status_s == cp_model.OPTIMAL or status_s == cp_model.FEASIBLE:
                                    st.success("✅ Optimisation Successful!")
                                else:
                                    st.warning("⚠️ Standby Pass Unsuccessful")
                            else:
                                if status_s == cp_model.OPTIMAL or status_s == cp_model.FEASIBLE:
                                    st.warning("⚠️ Duty Pass Unsuccessful")
                                else:
                                    st.warning("⚠️ No Solution Found")
                except Exception:
                    st.error("❌ Critical Error Detected")
                    st.code(traceback.format_exc())

        # planning buttons

        final_name = st.session_state.get('active_sh_name', spreadsheet_name)

        col1, col2 = st.columns([2,3])

        with col1:
            st.write("Step 2:")
        
        with col2:
            if st.button("💾 Save the Optimisation"):
                if 'planned_df' in st.session_state:
                    # 1. archive the original MMYYC using personal Drive account
                    with st.spinner("🧳 Creating Archive..."):
                        personal_drive = get_personal_drive_service()
                        folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
                        planner_engine.archive_source_sheet(client, final_name, mmyy, folder_id, personal_drive)
                    
                    # 2. write output
                    with st.spinner("✏️ Writing Output..."):
                        planner_engine.create_backup_and_output(
                            client, final_name, mmyy,
                            st.session_state['planned_df'],
                            st.session_state['n_scale'],
                            st.session_state['ranges']
                        )

                    # 3. create next month template
                    with st.spinner("⏭️ Preparing Next Month..."):
                        _, next_file_name = planner_engine.generate_next_month_template(
                            client, final_name, mmyy,
                            st.session_state['planned_df'],
                            st.session_state['ranges']
                        )

                        st.success(f"✅ Done!")
                        st.session_state['last_saved_mmyy'] = mmyy
                        st.session_state.pop('planned_df', None)
                else:
                    st.warning("⚠️ Run the optimiser first!")

        # undo button — only shown after a save has been performed this session
        if st.session_state.get('last_saved_mmyy'):
            saved_mmyy = st.session_state['last_saved_mmyy']

            # calculate next month label to know which template sheet to delete
            _s_m = int(saved_mmyy[:2])
            _s_y = int(saved_mmyy[2:])
            _next_m = _s_m + 1 if _s_m < 12 else 1
            _next_y = _s_y if _s_m < 12 else _s_y + 1
            _next_mmyy = f"{_next_m:02d}{_next_y:02d}"[2:] # MMYY format
            _next_mmyy = f"{_next_m:02d}{str(_next_y)[-2:]}"
            _d_sheet   = f"{saved_mmyy}D"
            _next_c    = f"{_next_mmyy}C"

            st.markdown("---")
            col_u1, col_u2 = st.columns([2, 3])
            with col_u1:
                st.write("Undo last save:")
            with col_u2:
                if st.button(f"↩️ Undo Save ({saved_mmyy})", type="secondary"):
                    st.session_state['confirm_undo'] = True

            if st.session_state.get('confirm_undo'):
                st.warning(
                    f"⚠️ This will delete **{_d_sheet}** and **{_next_c}** from the sheet. "
                    f"This cannot be undone. Are you sure?"
                )
                col_yes, col_no = st.columns(2)
                with col_yes:
                    if st.button("✅ Yes, Undo", use_container_width=True):
                        try:
                            sh_undo = client.open(final_name)
                            deleted = []
                            for sheet_name in [_d_sheet, _next_c]:
                                try:
                                    ws_del = sh_undo.worksheet(sheet_name)
                                    sh_undo.del_worksheet(ws_del)
                                    deleted.append(sheet_name)
                                except gspread.exceptions.WorksheetNotFound:
                                    pass  # already gone, not an error
                            fetch_sheet_data.clear()
                            st.session_state.pop('last_saved_mmyy', None)
                            st.session_state.pop('confirm_undo', None)
                            if deleted:
                                st.success(f"✅ Deleted: {', '.join(deleted)}")
                            else:
                                st.info("ℹ️ No sheets found to delete — they may have already been removed.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Undo failed: {e}")
                            st.code(traceback.format_exc())
                with col_no:
                    if st.button("❌ Cancel", use_container_width=True):
                        st.session_state.pop('confirm_undo', None)
                        st.rerun()

        # --------------------------------------------------
        # ADD PERSONNEL
        # --------------------------------------------------

        st.markdown("---")
        st.subheader("👤 Add Personnel")

        with st.container(border=True):
            col_name, col_branch = st.columns(2)
            with col_name:
                new_name = st.text_input("Full Name (as it should appear)", key="new_person_name").strip().upper()
            with col_branch:
                new_branch = st.text_input("Branch (e.g. OS1)", key="new_person_branch").strip().upper()

            if st.button("➕ Add Person", use_container_width=True):
                if not new_name or not new_branch:
                    st.error("❌ Please enter both name and branch.")
                else:
                    # set a confirmation flag in session_state
                    st.session_state["confirm_add_person"] = True
            if st.session_state.get("confirm_add_person"):
                if st.button(f"⚠️ Confirm Add {new_name} ({new_branch})"):
                    try:
                        sh = convert_if_excel(client, spreadsheet_name)
                        year_str = str(2000 + int(mmyy[2:]))

                        def find_insert_row(rows, name_col, branch_col, branch, new_name, data_start_row):
                            branch_rows = []
                            for i, r in enumerate(rows):
                                b = r[branch_col].strip().upper() if len(r) > branch_col else ""
                                n = r[name_col].strip().upper() if len(r) > name_col else ""
                                if b == branch and n:
                                    branch_rows.append((i + data_start_row, n))
                            if not branch_rows:
                                for i, r in enumerate(rows):
                                    n = r[name_col].strip() if len(r) > name_col else ""
                                    if not n:
                                        return i + data_start_row
                                return len(rows) + data_start_row
                            for gs_row, existing_name in branch_rows:
                                if new_name < existing_name:
                                    return gs_row
                            return branch_rows[-1][0] + 1

                        def insert_row_and_copy(sh, sheet_name, insert_gs_row, template_gs_row):
                            ws = sh.worksheet(sheet_name)
                            sheet_id = ws.id
                            body = {"requests": [
                                {"insertDimension": {
                                    "range": {
                                        "sheetId": sheet_id,
                                        "dimension": "ROWS",
                                        "startIndex": insert_gs_row - 1,
                                        "endIndex": insert_gs_row
                                    },
                                    "inheritFromBefore": False
                                }},
                                {"copyPaste": {
                                    "source": {
                                        "sheetId": sheet_id,
                                        "startRowIndex": template_gs_row - 1,
                                        "endRowIndex": template_gs_row,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 70
                                    },
                                    "destination": {
                                        "sheetId": sheet_id,
                                        "startRowIndex": insert_gs_row - 1,
                                        "endRowIndex": insert_gs_row,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 70
                                    },
                                    "pasteType": "PASTE_FORMULA",
                                    "pasteOrientation": "NORMAL"
                                }}
                            ]}
                            sh.batch_update(body)
                            return sh.worksheet(sheet_name)

                        # 1. NAMELIST
                        with st.spinner("📋 Updating Namelist..."):
                            nl_data = sh.worksheet("Namelist").get_all_values()
                            nl_rows = nl_data[1:]
                            nl_insert = find_insert_row(nl_rows, 1, 2, new_branch, new_name, data_start_row=2)
                            nl_ws = insert_row_and_copy(sh, "Namelist", nl_insert, template_gs_row=2)
                            nl_ws.update(f'B{nl_insert}', [[new_name]])
                            nl_ws.update(f'C{nl_insert}', [[new_branch]])
                            nl_ws.update(f'D{nl_insert}', [['NON-DRIVER']])

                        # 2. YEAR SHEET
                        with st.spinner(f"📅 Updating {year_str} sheet..."):
                            yr_data = sh.worksheet(year_str).get_all_values()
                            yr_rows = yr_data[2:]
                            yr_insert = find_insert_row(yr_rows, 1, 2, new_branch, new_name, data_start_row=3)
                            yr_ws = insert_row_and_copy(sh, year_str, yr_insert, template_gs_row=3)
                            yr_ws.update(f'B{yr_insert}', [[new_name]])

                        # 3. C SHEET
                        with st.spinner(f"📄 Updating {mmyy}C sheet..."):
                            c_sheet = f"{mmyy}C"
                            c_data = sh.worksheet(c_sheet).get_all_values()
                            c_rows = c_data[3:]
                            c_insert = find_insert_row(c_rows, 1, 2, new_branch, new_name, data_start_row=4)
                            c_ws = insert_row_and_copy(sh, c_sheet, c_insert, template_gs_row=4)
                            c_ws.update(f'B{c_insert}', [[new_name]])
                            c_ws.batch_clear([f"E{c_insert}:AI{c_insert}"])
                            c_ws.update(f'AQ{c_insert}', [[0.00]])
                            c_ws.update(f'AR{c_insert}', [['NEW']])

                        # clear caches
                        fetch_sheet_data.clear()
                        for key in list(st.session_state.keys()):
                            if key.startswith("roster_") or key.startswith("adj_data_"):
                                st.session_state.pop(key)

                        st.success(f"✅ {new_name} ({new_branch}) added to Namelist, {year_str}, and {mmyy}C!")

                    except Exception as e:
                        st.error(f"❌ Failed to add person: {e}")
                        st.code(traceback.format_exc())
        
        st.subheader("🗑️ Remove Personnel")

        with st.container(border=True):
            names_for_removal = fetch_namelist(client, spreadsheet_name)
            remove_name = st.selectbox("Select Person to Remove", options=[""] + names_for_removal, key="remove_person_name")

            if st.button("🗑️ Remove Person", use_container_width=True):
                if not remove_name:
                    st.error("❌ Please select a person.")
                else:
                    # set a confirmation flag in session_state
                    st.session_state["confirm_remove_person"] = remove_name
            if st.session_state.get("confirm_remove_person"):
                if st.button(f"⚠️ Confirm Remove {st.session_state['confirm_remove_person']}"):
                    try:
                        sh = convert_if_excel(client, spreadsheet_name)
                        c_sheet = f"{mmyy}C"
                        c_ws = sh.worksheet(c_sheet)

                        # find the person's row
                        cell = c_ws.find(remove_name, in_column=2)
                        if not cell:
                            st.error(f"❌ {remove_name} not found in {c_sheet}.")
                        else:
                            # delete the row entirely
                            c_ws.delete_rows(cell.row)

                            # clear caches
                            fetch_sheet_data.clear()
                            for key in list(st.session_state.keys()):
                                if key.startswith("roster_") or key.startswith("adj_data_"):
                                    st.session_state.pop(key)

                            st.success(f"✅ {remove_name} removed from {c_sheet}.")

                    except Exception as e:
                        st.error(f"❌ Failed to remove person: {e}")
                        st.code(traceback.format_exc())

        # --------------------------------------------------
        # HOLIDAY SECTION
        # --------------------------------------------------
        st.markdown("---")
        st.subheader("🗓️ Holiday Management")

        hol_tab1, hol_tab2 = st.tabs(["📥 Upload Holidays", "👥 Assign Duty"])

        # ── Tab 1: Upload ICS ──
        with hol_tab1:
            uploaded_ics = st.file_uploader("Upload .ics file", type=["ics"], key="ics_uploader")

            if uploaded_ics:
                try:
                    from datetime import date as _date, timedelta as _td
                    content_ics = uploaded_ics.read().decode("utf-8")
                    lines = content_ics.splitlines()

                    holidays = []
                    current_date = None
                    current_name = None
                    for line in lines:
                        line = line.strip()
                        if line.startswith("DTSTART"):
                            date_str = line.split(":")[-1].strip()
                            try:
                                current_date = _date(int(date_str[:4]), int(date_str[4:6]), int(date_str[6:8]))
                            except:
                                current_date = None
                        elif line.startswith("SUMMARY"):
                            current_name = line.split(":", 1)[-1].strip()
                        elif line == "END:VEVENT":
                            if current_date and current_name:
                                holidays.append((current_date, current_name))
                                if current_date.weekday() == 6:
                                    holidays.append((current_date + _td(days=1), f"{current_name} (in lieu)"))
                            current_date = None
                            current_name = None

                    if holidays:
                        yr = holidays[0][0].year
                        holidays.append((_date(yr, 12, 31), "NEW YEAR'S EVE"))
                        holidays.append((_date(yr, 12, 24), "CHRISTMAS EVE"))
                        cny_names = ["chinese new year", "cny"]
                        cny_dates = sorted([h[0] for h in holidays if any(c in h[1].lower() for c in cny_names)])
                        if cny_dates:
                            holidays.append((cny_dates[0] - _td(days=1), "CHINESE NEW YEAR EVE"))

                    holidays.sort(key=lambda h: h[0])

                    preview_df = pd.DataFrame([
                        {"Date": h[0].strftime("%-d %b %Y"), "Day": h[0].strftime("%a"), "Holiday": h[1].upper()}
                        for h in holidays
                    ])
                    st.dataframe(preview_df, use_container_width=True, hide_index=True)
                    st.caption(f"{len(holidays)} holidays (including eves and in lieu days)")

                    if st.button("📥 Write to Holiday Sheet", use_container_width=True, key="write_holidays"):
                        try:
                            sh_hol = convert_if_excel(client, spreadsheet_name)
                            hol_ws = sh_hol.worksheet("Holiday")
                            hol_data = hol_ws.get_all_values()
                            last_data_row = len(hol_data)
                            start_row = last_data_row + 1
                            end_row = start_row + len(holidays) - 1
                            sheet_id = hol_ws.id
                            n = len(holidays)

                            sh_hol.batch_update({"requests": [{
                                "insertDimension": {
                                    "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                              "startIndex": last_data_row, "endIndex": last_data_row + n},
                                    "inheritFromBefore": True
                                }
                            }]})
                            sh_hol.batch_update({"requests": [{
                                "copyPaste": {
                                    "source": {"sheetId": sheet_id,
                                               "startRowIndex": last_data_row - 1, "endRowIndex": last_data_row,
                                               "startColumnIndex": 0, "endColumnIndex": 5},
                                    "destination": {"sheetId": sheet_id,
                                                    "startRowIndex": last_data_row, "endRowIndex": last_data_row + n,
                                                    "startColumnIndex": 0, "endColumnIndex": 5},
                                    "pasteType": "PASTE_FORMAT", "pasteOrientation": "NORMAL"
                                }
                            }]})

                            rows_to_write = [
                                [h[1].upper(), h[0].strftime("%-d %b %Y"), h[0].strftime("%a"), "", ""]
                                for h in holidays
                            ]
                            hol_ws.update(f"A{start_row}:E{end_row}", rows_to_write)
                            fetch_sheet_data.clear()
                            st.success(f"✅ Written {n} holidays to Holiday sheet!")

                        except Exception as e:
                            st.error(f"❌ Failed to write holidays: {e}")
                            st.code(traceback.format_exc())

                except Exception as e:
                    st.error(f"❌ Failed to parse ICS file: {e}")

        # ── Tab 2: Assign Duty ──
        with hol_tab2:
            try:
                # load holiday sheet and namelist
                hol_raw = fetch_sheet_data(client, spreadsheet_name, "Holiday")
                all_names = fetch_namelist(client, spreadsheet_name)

                if not hol_raw or len(hol_raw) < 2:
                    st.caption("No holidays found in Holiday sheet.")
                else:
                    # identify female names (contain "(F)")
                    female_names = [n for n in all_names if "(F)" in n.upper()]

                    # parse holiday rows: col A=name, B=date, C=day, D=name1, E=name2
                    # filter to selected year from mmyy
                    target_year = 2000 + int(mmyy[2:])

                    hol_rows = []
                    for i, row in enumerate(hol_raw[1:], start=2):  # skip header, track sheet row
                        if not row or not row[0].strip():
                            continue
                        hol_name = row[0].strip()
                        hol_date_str = row[1].strip() if len(row) > 1 else ""
                        hol_day = row[2].strip() if len(row) > 2 else ""
                        existing_n1 = row[3].strip() if len(row) > 3 else ""
                        existing_n2 = row[4].strip() if len(row) > 4 else ""

                        # parse date to check year
                        hol_date = None
                        for fmt in ["%d %b %Y", "%-d %b %Y", "%Y-%m-%d"]:
                            try:
                                from datetime import datetime as _dt
                                hol_date = _dt.strptime(hol_date_str, fmt).date()
                                break
                            except:
                                continue

                        if hol_date and hol_date.year == target_year:
                            hol_rows.append({
                                "sheet_row": i,
                                "name": hol_name,
                                "date": hol_date,
                                "date_str": hol_date_str,
                                "day": hol_day,
                                "n1": existing_n1,
                                "n2": existing_n2,
                            })

                    if not hol_rows:
                        st.caption(f"No holidays found for {target_year}. Upload holidays first.")
                    else:
                        # build "did holiday last year" lookup from same sheet
                        # look for rows with same month/day but previous year
                        last_year = target_year - 1
                        last_year_workers = set()
                        for row in hol_raw[1:]:
                            if not row or not row[0].strip():
                                continue
                            hol_date_str_ly = row[1].strip() if len(row) > 1 else ""
                            n1_ly = row[3].strip() if len(row) > 3 else ""
                            n2_ly = row[4].strip() if len(row) > 4 else ""
                            hol_date_ly = None
                            for fmt in ["%d %b %Y", "%-d %b %Y", "%Y-%m-%d"]:
                                try:
                                    from datetime import datetime as _dt
                                    hol_date_ly = _dt.strptime(hol_date_str_ly, fmt).date()
                                    break
                                except:
                                    continue
                            if hol_date_ly and hol_date_ly.year == last_year:
                                if n1_ly: last_year_workers.add(n1_ly.upper())
                                if n2_ly: last_year_workers.add(n2_ly.upper())

                        st.caption(f"Showing {len(hol_rows)} holidays for {target_year}. "
                                   f"Excluded from random (did duty last year): {len(last_year_workers)} people.")

                        # build input table
                        name_options = [""] + all_names
                        assignments = {}

                        st.markdown("**Enter names for each holiday (leave blank for random assignment):**")
                        for h in hol_rows:
                            col_hol, col_n1, col_n2 = st.columns([3, 2, 2])
                            with col_hol:
                                st.markdown(f"**{h['name']}**")
                                st.caption(f"{h['date_str']} ({h['day']})")
                            with col_n1:
                                n1 = st.selectbox("Name 1", name_options,
                                                  index=name_options.index(h['n1']) if h['n1'] in name_options else 0,
                                                  key=f"hol_n1_{h['sheet_row']}", label_visibility="collapsed")
                            with col_n2:
                                n2 = st.selectbox("Name 2", name_options,
                                                  index=name_options.index(h['n2']) if h['n2'] in name_options else 0,
                                                  key=f"hol_n2_{h['sheet_row']}", label_visibility="collapsed")
                            assignments[h['sheet_row']] = {"h": h, "n1": n1, "n2": n2}
                            st.markdown("<hr style='margin:4px 0;border:none;border-top:1px solid rgba(255,255,255,0.1);'>", unsafe_allow_html=True)

                        if st.button("🎲 Generate & Save", use_container_width=True, key="gen_hol_duty"):
                            import random as _random

                            # eligible pool: exclude last year workers
                            eligible = [n for n in all_names if n.upper() not in last_year_workers]

                            updates = []
                            for sr, asgn in assignments.items():
                                h = asgn["h"]
                                n1 = asgn["n1"].strip()
                                n2 = asgn["n2"].strip()

                                # determine if any female is manually assigned
                                n1_is_female = n1 and "(F)" in n1.upper()
                                n2_is_female = n2 and "(F)" in n2.upper()

                                # fill blanks randomly
                                used = {n1.upper(), n2.upper()} - {""}
                                pool = [n for n in eligible if n.upper() not in used]

                                if not n1 and not n2:
                                    # both blank — check if we should assign females
                                    # randomly decide: if females available and not excluded, can assign both
                                    female_pool = [n for n in female_names if n.upper() not in last_year_workers]
                                    if female_pool and len(female_pool) >= 2 and _random.random() < 0.3:
                                        chosen = _random.sample(female_pool, 2)
                                        n1, n2 = chosen[0], chosen[1]
                                    else:
                                        non_female_pool = [n for n in pool if "(F)" not in n.upper()]
                                        if len(non_female_pool) >= 2:
                                            chosen = _random.sample(non_female_pool, 2)
                                            n1, n2 = chosen[0], chosen[1]
                                        elif len(pool) >= 2:
                                            chosen = _random.sample(pool, 2)
                                            n1, n2 = chosen[0], chosen[1]

                                elif not n1:
                                    # n2 is set — if n2 is female, n1 must be female
                                    if n2_is_female:
                                        female_pool = [n for n in female_names
                                                       if n.upper() not in last_year_workers
                                                       and n.upper() != n2.upper()]
                                        n1 = _random.choice(female_pool) if female_pool else ""
                                    else:
                                        non_fem = [n for n in pool if "(F)" not in n.upper()]
                                        n1 = _random.choice(non_fem) if non_fem else (_random.choice(pool) if pool else "")

                                elif not n2:
                                    # n1 is set — if n1 is female, n2 must be female
                                    if n1_is_female:
                                        female_pool = [n for n in female_names
                                                       if n.upper() not in last_year_workers
                                                       and n.upper() != n1.upper()]
                                        n2 = _random.choice(female_pool) if female_pool else ""
                                    else:
                                        used2 = {n1.upper()}
                                        non_fem = [n for n in eligible
                                                   if n.upper() not in used2 and "(F)" not in n.upper()]
                                        n2 = _random.choice(non_fem) if non_fem else ""

                                updates.append({"sheet_row": sr, "n1": n1, "n2": n2})

                            # write to sheet
                            try:
                                sh_hol2 = convert_if_excel(client, spreadsheet_name)
                                hol_ws2 = sh_hol2.worksheet("Holiday")
                                batch = []
                                for u in updates:
                                    batch.append({"range": f"D{u['sheet_row']}:E{u['sheet_row']}",
                                                  "values": [[u['n1'], u['n2']]]})
                                hol_ws2.batch_update(batch, value_input_option='USER_ENTERED')
                                fetch_sheet_data.clear()
                                st.success("✅ Holiday duties saved!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to save: {e}")

            except Exception as e:
                st.error(f"❌ Could not load holiday duty section: {e}")

    if admin_page == "✏️ Editing":

        _now = date.today(); _opts = [f"{m:02d}{str(y)[2:]}" for y in [_now.year, _now.year+1] for m in range(1,13)]
        mmyy = st.selectbox("Month/Year (MMYY) to edit", options=_opts, key="edit_mmyy")
        spreadsheet_name = "MASTER SHEET"

        curr_m, curr_y = int(mmyy[:2]), int(mmyy[2:])
        m_old = curr_m - 1
        if curr_m == 1:
            m_old = 12
            y_old = curr_y - 1
        else:
            y_old = curr_y

        # point allocations — use planning page slider values if available, else defaults
        point_allocations = {
            "weekday_points": st.session_state.get("weekday_slider", 1.0),
            "friday_points": st.session_state.get("friday_slider", 1.0),
            "weekend_points": st.session_state.get("weekend_slider", 2.0),
            "holiday_points": st.session_state.get("holiday_slider", 2.0),
        }

        st.info(f"Editing **{mmyy}**!")

        try:
            personal_drive = get_personal_drive_service()
            folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
            sheet_id = fetch_spreadsheet_id(personal_drive, folder_id, spreadsheet_name)
            if sheet_id:
                st.success(f"✅ Connected to storage!")
            else:
                st.warning(f"⚠️ Connection error: storage failed!")
        except Exception as e:
            st.error(f"❌ Storage failed: {e}")

        roster_cache_key = f"roster_{mmyy}"
        if roster_cache_key not in st.session_state:
            roster_data, sheet_used, err = user_engine.calendar_view(client, spreadsheet_name, mmyy)
            if not err:
                st.session_state[roster_cache_key] = {"roster_data": roster_data, "sheet_used": sheet_used, "err": err}
        else:
            cached_r = st.session_state[roster_cache_key]
            roster_data, sheet_used, err = cached_r["roster_data"], cached_r["sheet_used"], cached_r["err"]

        sh = client.open(spreadsheet_name)
        
        if err:
            st.warning(f"⚠️ Roster not yet finalised or accessible: {err}")
        else:

            if sheet_used == "D":
                st.success("✅ Showing finalised roster")
            elif sheet_used == "C":
                st.info("ℹ️ Showing draft constraints — roster not yet finalised")

            # segmented control: Calendar / C Sheet / Summary
            cal_view_mode = st.segmented_control(
                "View", options=["Calendar", "C Sheet", "Summary"], default="Calendar", key="edit_cal_mode"
            )

            if cal_view_mode == 'Summary':
                year_str_sum = str(2000 + int(mmyy[2:]))
                try:
                    yr_cache_key = f'yr_summary_{year_str_sum}'
                    if yr_cache_key not in st.session_state:
                        yr_raw = fetch_sheet_data(client, spreadsheet_name, year_str_sum)
                        st.session_state[yr_cache_key] = yr_raw
                    yr_raw = st.session_state[yr_cache_key]
                    if yr_raw:
                        yr_df = pd.DataFrame(yr_raw)
                        st.dataframe(yr_df, use_container_width=True, hide_index=True)
                    else:
                        st.warning(f'⚠️ No data found in {year_str_sum} sheet.')
                except Exception as e:
                    st.error(f'❌ Could not load year sheet: {e}')

            if cal_view_mode == 'C Sheet':
                c_sheet_name = f"{mmyy}C"
                c_cache_key = f"c_sheet_preview_{mmyy}"
                try:
                    if c_cache_key not in st.session_state:
                        raw_c = fetch_sheet_data(client, spreadsheet_name, c_sheet_name)
                        st.session_state[c_cache_key] = raw_c
                    raw_c = st.session_state[c_cache_key]

                    if not raw_c or len(raw_c) < 2:
                        st.warning(f"⚠️ No data found in {c_sheet_name}.")
                    else:
                        # Row 3 (index 2) is the header; data starts row 4 (index 3) onwards.
                        # Columns A–AR = indices 0–43.
                        header_row = raw_c[2][:44]
                        data_rows  = raw_c[2:]

                        # Find last row with a name in column B (index 1)
                        last_name_idx = 0
                        for i, row in enumerate(data_rows):
                            name_val = row[1].strip() if len(row) > 1 else ""
                            if name_val:
                                last_name_idx = i
                        data_rows = data_rows[:last_name_idx + 1]

                        # Trim each row to 44 columns (A–AR), padding short rows
                        trimmed = []
                        for row in data_rows:
                            padded = (row + [""] * 44)[:44]
                            trimmed.append(padded)

                        c_df = pd.DataFrame(trimmed, columns=header_row)
                        st.dataframe(c_df, use_container_width=True, hide_index=True)
                        st.caption(f"Showing {c_sheet_name} — columns A to AR, {len(trimmed)} person(s).")
                except Exception as e:
                    st.error(f"❌ Could not load {c_sheet_name}: {e}")

            if cal_view_mode == 'Calendar':

                # 1. Date Math & Setup
                first_day = date(curr_y, curr_m, 1)
                start_padding = (first_day.weekday()) % 7 
                num_days = calendar.monthrange(curr_y, curr_m)[1]
                days_of_week = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

                # 2. Hybrid CSS: Pill for Duty, Plain for Standby
                st.markdown("""
                    <style>
                        @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400;600;700&display=swap');

                        .cal-container {
                            border: 1px solid #444;
                            border-radius: 15px;
                            overflow: hidden; 
                            margin-top: 10px;
                            font-family: 'Source Sans Pro', sans-serif !important;
                        }
                    
                        .cal-table { 
                            width: 100%; 
                            border-collapse: collapse; 
                            table-layout: fixed;
                            background-color: #262730; /* Darker background to make white text/blue pills pop */
                        }

                        .cal-th { 
                            background-color: #000000; 
                            color: #ffffff; 
                            padding: 12px; 
                            text-align: center; 
                            border-bottom: 1px solid #444;
                            font-weight: 600 !important;
                        }

                        .cal-td { 
                            vertical-align: top; 
                            border: 0.5px solid rgba(255, 255, 255, 0.1); 
                            height: 125px; 
                            padding: 10px; 
                        }

                        .day-num { 
                            font-weight: 700 !important; 
                            font-size: 1rem; 
                            margin-bottom: 10px; 
                            display: block;
                            color: #ffffff;
                        }

                        /* 🚨 DUTY: The "Pill" Style */
                        .duty-item { 
                            font-family: 'Source Sans Pro', sans-serif !important;
                            font-size: 10px; 
                            line-height: 1.2; 
                            margin-bottom: 6px; 
                            font-weight: 600 !important;
                            white-space: nowrap;
                            overflow: hidden;
                            text-overflow: ellipsis;
                            display: block;
                            padding: 4px 8px;
                            border-radius: 6px;
                            background-color: #007bff; 
                            border: 1px solid #0056b3;
                            color: white !important;
                        }
                    
                        /* ⏳ STANDBY: Plain White Text Style */
                        .standby-item { 
                            font-family: 'Source Sans Pro', sans-serif !important;
                            font-size: 11px; 
                            line-height: 1.4; 
                            margin-bottom: 3px; 
                            font-weight: 400 !important;
                            white-space: nowrap;
                            overflow: hidden;
                            text-overflow: ellipsis;
                            display: block;
                            color: white !important; /* Plain white text, no background */
                            padding-left: 2px;
                        }
                    </style>
                """, unsafe_allow_html=True)

                # 3. Build Table
                html_table = '<div class="cal-container"><table class="cal-table"><thead><tr>'
                for day_name in days_of_week:
                    html_table += f'<th class="cal-th">{day_name}</th>'
                html_table += '</tr></thead><tbody><tr>'

                for i in range(start_padding):
                    html_table += '<td class="cal-td"></td>'

                current_col = start_padding

                for day in range(1, num_days + 1):
                    if current_col == 7:
                        html_table += '</tr><tr>'
                        current_col = 0
                
                    day_info = roster_data.get(str(day), {"duty": [], "standby": []})
                
                    cell_content = f'<span class="day-num">{day}</span>'
                
                    # Duty names get the blue pill
                    for d_name in day_info["duty"]:
                        cell_content += f'<div class="duty-item" title="Duty: {d_name}">{d_name}</div>'
                
                    # Standby names get plain white text
                    for s_name in day_info["standby"]:
                        cell_content += f'<div class="standby-item" title="Standby: {s_name}">{s_name}</div>'
                
                    html_table += f'<td class="cal-td">{cell_content}</td>'
                    current_col += 1

                while current_col < 7:
                    html_table += '<td class="cal-td"></td>'
                    current_col += 1

                html_table += '</tr></tbody></table></div>'

                st.markdown(html_table, unsafe_allow_html=True)
    
        # --------------------------------------------------
        # MANUAL ADJUSTMENTS TABLE
        # --------------------------------------------------
        st.markdown("### 🔄 Manual Adjustments")
        adj_cache_key = f"adj_data_{mmyy}"
        try:
            # use cached D sheet data if available, otherwise fetch
            if adj_cache_key in st.session_state:
                adj_raw = st.session_state[adj_cache_key]["raw_d_data"]
            else:
                adj_raw = fetch_sheet_data(client, spreadsheet_name, f"{mmyy}D")

            # find "NAME 1" header row in col AU (index 46)
            header_row_idx = None
            for li, rrow in enumerate(adj_raw):
                if len(rrow) > 46 and rrow[46].strip().upper() == "NAME 1":
                    header_row_idx = li
                    break

            if header_row_idx is None:
                st.caption("No adjustments table found in sheet.")
            else:
                # collect rows below header until empty
                adj_entries = []
                for rrow in adj_raw[header_row_idx + 1:]:
                    name1 = rrow[46].strip() if len(rrow) > 46 else ""
                    if not name1:
                        break
                    name2 = rrow[47].strip() if len(rrow) > 47 else ""
                    day   = rrow[48].strip() if len(rrow) > 48 else ""
                    dtype = rrow[49].strip() if len(rrow) > 49 else ""
                    adj_entries.append({"Name 1": name1, "Name 2": name2, "Day": day, "Day Type": dtype})

                if adj_entries:
                    st.dataframe(pd.DataFrame(adj_entries), use_container_width=True, hide_index=True)
                else:
                    st.caption("No adjustments recorded yet.")
        except Exception as e:
            st.caption(f"Could not load adjustments: {e}")

        # --------------------------------------------------
        # DOWNLOADS
        # --------------------------------------------------
        st.markdown("### 📥 Downloads")
        dl_col1, dl_col2, dl_col3 = st.columns(3)

        # Button 1: Calendar PDF
        with dl_col1:
            if sheet_used == "D" and roster_data:
                if st.button("🗓️ Calendar PDF", use_container_width=True):
                    try:
                        cal_first_day = date(2000 + int(mmyy[2:]), int(mmyy[:2]), 1)
                        cal_num_days = calendar.monthrange(2000 + int(mmyy[2:]), int(mmyy[:2]))[1]
                        cal_start_pad = cal_first_day.weekday()
                        month_label = cal_first_day.strftime("%B %Y")
                        dow = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

                        css = (
                            "@page { size: A4 landscape; margin: 1cm; }"
                            "body { font-family: Helvetica, Arial, sans-serif; background: white; color: black; margin: 0; }"
                            "h2 { text-align: center; margin-bottom: 6px; font-size: 15px; }"
                            
                            # Force the table to allow overflow instead of shrinking text
                            "table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 2px solid #555; -pdf-keep-in-frame-mode: overflow; }"
                            
                            "th { background-color: #333; color: white; padding: 5px 0; text-align: center; font-size: 13px; font-weight: bold; }"
                            "td { border: 1px solid #aaa; vertical-align: top; height: 100px; padding: 2px; width: 14.28%; -pdf-keep-in-frame-mode: overflow; }"
                            
                            ".day-num { font-weight: bold; font-size: 15px; display: block; margin: 2px 0 2px 1px; color: #111; }"
                            
                            # We add -pdf-keep-in-frame-mode here as well and nowrap to ensure font-size is respected
                            ".duty-item { font-size: 12px; background-color: #1a73e8; color: white; border-radius: 3px; padding: 1px 3px; display: block; font-weight: bold; line-height: 1.2; margin: 0 1px 1px 1px; white-space: nowrap; -pdf-keep-in-frame-mode: overflow; }"
                            ".standby-item { font-size: 12px; color: #222; display: block; line-height: 1.2; margin: 0 1px 0 1px; white-space: nowrap; -pdf-keep-in-frame-mode: overflow; }"
                        )

                        cal_html = "<!DOCTYPE html><html><head><meta charset='utf-8'><style>" + css + "</style></head><body>"
                        cal_html += "<h2>" + month_label + " Duty Roster</h2>"
                        cal_html += "<table><thead><tr>"
                        for d in dow:
                            cal_html += "<th>" + d + "</th>"
                        cal_html += "</tr></thead><tbody><tr>"

                        for _ in range(cal_start_pad):
                            cal_html += "<td></td>"
                        col_idx = cal_start_pad

                        for day in range(1, cal_num_days + 1):
                            if col_idx == 7:
                                cal_html += "</tr><tr>"
                                col_idx = 0
                            info = roster_data.get(str(day), {"duty": [], "standby": []})
                            cal_html += "<td><span class='day-num'>" + str(day) + "</span>"
                            for n in info["duty"]:
                                # truncate long names to avoid overflow
                                display_name = n if len(n) <= 20 else n[:19] + "."
                                cal_html += "<span class='duty-item'>" + display_name + "</span>"
                            for n in info["standby"]:
                                display_name = n if len(n) <= 20 else n[:19] + "."
                                cal_html += "<span class='standby-item'>" + display_name + "</span>"
                            cal_html += "</td>"
                            col_idx += 1

                        while col_idx < 7:
                            cal_html += "<td></td>"
                            col_idx += 1
                        cal_html += "</tr></tbody></table></body></html>"

                        from xhtml2pdf import pisa
                        import io
                        pdf_buffer = io.BytesIO()
                        pisa.CreatePDF(cal_html, dest=pdf_buffer)
                        pdf_bytes = pdf_buffer.getvalue()
                        st.download_button(
                            label="⬇️ Download Calendar PDF",
                            data=pdf_bytes,
                            file_name=mmyy + "_roster_calendar.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error("❌ Calendar PDF failed: " + str(e))
            else:
                st.button("🗓️ Calendar PDF", disabled=True, use_container_width=True,
                          help="Only available when D sheet exists")

        # Button 2: D Sheet PDF via Google Sheets export API
        with dl_col2:
            if sheet_used == "D":
                if st.button("📄 D Sheet PDF", use_container_width=True):
                    try:
                        personal_drive_dl = get_personal_drive_service()
                        folder_id_dl = st.secrets["app_config"]["personal_drive_folder_id"]
                        sid_dl = fetch_spreadsheet_id(personal_drive_dl, folder_id_dl, spreadsheet_name)
                        if not sid_dl:
                            st.error("❌ Could not find spreadsheet.")
                        else:
                            sh_dl = client.open_by_key(sid_dl)
                            d_ws_dl = sh_dl.worksheet(mmyy + "D")
                            sheet_gid_dl = d_ws_dl.id

                            d_raw_dl = fetch_sheet_data(client, spreadsheet_name, mmyy + "D")
                            last_name_row_dl = 3
                            for ri, rrow in enumerate(d_raw_dl[3:], start=4):
                                if len(rrow) > 1 and rrow[1].strip():
                                    last_name_row_dl = ri

                            export_url_dl = (
                                f"https://docs.google.com/spreadsheets/d/{sid_dl}/export"
                                f"?format=pdf"
                                f"&gid={sheet_gid_dl}"
                                f"&portrait=false"
                                f"&scale=4"
                                f"&gridlines=true"
                                f"&r1=0&c1=0&r2={last_name_row_dl}&c2=45"
                                f"&ir=false&ic=false"
                            )

                            info_dl = st.secrets["personal_account"]
                            creds_dl = Credentials(
                                token=info_dl["token"],
                                refresh_token=info_dl["refresh_token"],
                                client_id=info_dl["client_id"],
                                client_secret=info_dl["client_secret"],
                                token_uri=info_dl["token_uri"]
                            )
                            if creds_dl.expired:
                                creds_dl.refresh(Request())

                            resp_dl = http_requests.get(
                                export_url_dl,
                                headers={"Authorization": "Bearer " + creds_dl.token}
                            )
                            if resp_dl.status_code == 200:
                                st.download_button(
                                    label="⬇️ Download D Sheet PDF",
                                    data=resp_dl.content,
                                    file_name=mmyy + "D_sheet.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                            else:
                                st.error("❌ Export failed: HTTP " + str(resp_dl.status_code))
                    except Exception as e:
                        st.error("❌ D Sheet PDF failed: " + str(e))
            else:
                st.button("📄 D Sheet PDF", disabled=True, use_container_width=True,
                          help="Only available when D sheet exists")

        # Button 3: Master Sheet Excel (always available)
        with dl_col3:
            if st.button("📊 Master Sheet Excel", use_container_width=True):
                try:
                    personal_drive_xl = get_personal_drive_service()
                    folder_id_xl = st.secrets["app_config"]["personal_drive_folder_id"]
                    sid_xl = fetch_spreadsheet_id(personal_drive_xl, folder_id_xl, spreadsheet_name)
                    if not sid_xl:
                        st.error("❌ Could not find spreadsheet.")
                    else:
                        export_url_xl = (
                            "https://docs.google.com/spreadsheets/d/" + sid_xl + "/export"
                            "?format=xlsx"
                        )
                        info_xl = st.secrets["personal_account"]
                        creds_xl = Credentials(
                            token=info_xl["token"],
                            refresh_token=info_xl["refresh_token"],
                            client_id=info_xl["client_id"],
                            client_secret=info_xl["client_secret"],
                            token_uri=info_xl["token_uri"]
                        )
                        if creds_xl.expired:
                            creds_xl.refresh(Request())

                        resp_xl = http_requests.get(
                            export_url_xl,
                            headers={"Authorization": "Bearer " + creds_xl.token}
                        )
                        if resp_xl.status_code == 200:
                            st.download_button(
                                label="⬇️ Download Excel",
                                data=resp_xl.content,
                                file_name=spreadsheet_name + ".xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ Export failed: HTTP " + str(resp_xl.status_code))
                except Exception as e:
                    st.error("❌ Excel download failed: " + str(e))

        # --------------------------------------------------
        # SIDEBAR: MANUAL DUTY SWAP
        # --------------------------------------------------
        st.sidebar.title("📅 Editing Settings")
        st.sidebar.subheader("🔄 Manual Duty Swap")

        target_sheet_name = f"{mmyy}D"
        cache_key = f"adj_data_{mmyy}"

        try:
            personal_drive = get_personal_drive_service()
            folder_id = st.secrets["app_config"]["personal_drive_folder_id"]

            # cached Drive file ID lookup
            sheet_id = fetch_spreadsheet_id(personal_drive, folder_id, spreadsheet_name)
            if not sheet_id:
                raise FileNotFoundError(f"Could not find '{spreadsheet_name}'")
            sh_admin = client.open_by_key(sheet_id)

            # load D sheet and holiday data once, cache in session state
            if cache_key not in st.session_state:
                with st.spinner("📥 Loading sheet data..."):
                    raw_d_data = fetch_sheet_data(client, spreadsheet_name, target_sheet_name)
                    hol_data = fetch_sheet_data(client, spreadsheet_name, "Holiday")
                    # AU3 = row index 2, col index 46
                    scale_raw = raw_d_data[2][46] if len(raw_d_data) > 2 and len(raw_d_data[2]) > 46 else None
                    st.session_state[cache_key] = {
                        "raw_d_data": raw_d_data,
                        "scale_raw": scale_raw,
                        "hol_data": hol_data
                    }

            if st.sidebar.button("🔄 Refresh Data", key="refresh_adj"):
                st.session_state.pop(cache_key, None)
                fetch_sheet_data.clear()
                st.rerun()

            cached = st.session_state[cache_key]
            raw_d_data = cached["raw_d_data"]
            scale_raw = cached["scale_raw"]
            hol_data = cached["hol_data"]
            d_rows = raw_d_data[3:]  # data from row 4

            # keep ws reference only for writes
            adj_ws = sh_admin.worksheet(target_sheet_name)

            # parse scale
            try:
                adj_scale = float(scale_raw) if scale_raw and float(scale_raw) > 0 else 1.0
            except:
                adj_scale = 1.0

            # parse holiday dates from cached data
            holiday_dates = set()
            from datetime import datetime as dt
            for hrow in hol_data[1:]:
                if len(hrow) > 1 and hrow[1].strip():
                    try:
                        hdate = dt.strptime(hrow[1].strip(), "%Y-%m-%d").date() if "-" in hrow[1] else None
                        if hdate:
                            holiday_dates.add(hdate)
                    except:
                        pass

            # step 1: day picker
            view_y = curr_y + 2000
            num_days_in_month = calendar.monthrange(curr_y, curr_m)[1]
            day_options = [date(view_y, curr_m, d) for d in range(1, num_days_in_month + 1)]
            day_labels = [d.strftime("%d %b %Y (%a)") for d in day_options]

            selected_label = st.sidebar.selectbox(
                "Select Day",
                options=day_labels,
                key="swap_date_picker"
            )

            swap_date = day_options[day_labels.index(selected_label)]
            selected_day = swap_date.day
            date_col_idx = 4 + selected_day - 1  # 0-indexed (col E = index 4)

            # determine day type automatically
            weekday_num = swap_date.weekday()
            if swap_date in holiday_dates:
                day_type_code, day_type_label = "H", "Holiday"
                day_points = point_allocations["holiday_points"]
            elif weekday_num == 4:
                day_type_code, day_type_label = "F", "Friday"
                day_points = point_allocations["friday_points"]
            elif weekday_num >= 5:
                day_type_code, day_type_label = "WE", "Weekend"
                day_points = point_allocations["weekend_points"]
            else:
                day_type_code, day_type_label = "WD", "Weekday"
                day_points = point_allocations["weekday_points"]

            st.sidebar.caption(f"📅 Day type: **{day_type_label}** ({day_points} pts)")

            # build filtered name lists from cached data
            names_with_d = []
            names_without_d = []
            for drow in d_rows:
                if not drow or len(drow) < 2 or not drow[1].strip():
                    continue
                name = drow[1].strip()
                cell_val = drow[date_col_idx].strip().upper() if date_col_idx < len(drow) else ""
                if cell_val == "D":
                    names_with_d.append(name)
                else:
                    names_without_d.append(name)

            # step 2: person 1 — must have D on selected day
            person_1 = st.sidebar.selectbox(
                "Person giving up duty",
                options=[""] + names_with_d,
                key="swap_person_1"
            )

            # step 3: person 2 — must NOT have D on selected day
            person_2_options = [n for n in names_without_d if n != person_1]
            person_2 = st.sidebar.selectbox(
                "Person taking over duty",
                options=[""] + person_2_options,
                key="swap_person_2"
            )

            # step 4: save button
            if st.sidebar.button("💾 Save Swap", use_container_width=True):
                if not person_1:
                    st.sidebar.error("❌ No one has a D on this day.")
                elif not person_2:
                    st.sidebar.error("❌ Please select a person to take over.")
                else:
                    try:
                        p1_row = next(
                            (i + 4 for i, r in enumerate(d_rows)
                             if r and len(r) > 1 and r[1].strip() == person_1), None
                        )
                        p2_row = next(
                            (i + 4 for i, r in enumerate(d_rows)
                             if r and len(r) > 1 and r[1].strip() == person_2), None
                        )

                        if not p1_row or not p2_row:
                            st.sidebar.error("❌ Could not locate one or both people in the sheet.")
                        else:
                            # col letter for selected day
                            day_gs_col = 5 + selected_day - 1
                            day_col_letter = gspread.utils.rowcol_to_a1(1, day_gs_col)[:-1]

                            # read offsets from cached data (AQ = col index 42 in 0-indexed)
                            p1_offset_raw = raw_d_data[p1_row - 1][42] if len(raw_d_data[p1_row - 1]) > 42 else None
                            p2_offset_raw = raw_d_data[p2_row - 1][42] if len(raw_d_data[p2_row - 1]) > 42 else None
                            p1_offset = float(p1_offset_raw) if p1_offset_raw else 0.0
                            p2_offset = float(p2_offset_raw) if p2_offset_raw else 0.0

                            # compute new offsets
                            p1_new_offset = round((p1_offset / adj_scale) - day_points, 4)
                            p2_new_offset = round((p2_offset / adj_scale) + day_points, 4)

                            # find row with "NAME 1" in col AU (index 46), then write below it
                            header_row_idx = None
                            for li, rrow in enumerate(raw_d_data):
                                if len(rrow) > 46 and rrow[46].strip().upper() == "NAME 1":
                                    header_row_idx = li
                                    break

                            if header_row_idx is None:
                                raise ValueError("Could not find 'NAME 1' header in column AU")

                            # find next empty row below the header
                            next_log_row = header_row_idx + 2  # start one row below header (1-indexed)
                            for li in range(header_row_idx + 1, len(raw_d_data)):
                                row_au = raw_d_data[li][46] if len(raw_d_data[li]) > 46 else ""
                                if not row_au.strip():
                                    next_log_row = li + 1  # convert to 1-indexed
                                    break
                            else:
                                next_log_row = len(raw_d_data) + 1

                            updates = [
                                {'range': f'{day_col_letter}{p1_row}', 'values': [['']]},
                                {'range': f'{day_col_letter}{p2_row}', 'values': [['D']]},
                                {'range': f'AQ{p1_row}', 'values': [[p1_new_offset]]},
                                {'range': f'AQ{p2_row}', 'values': [[p2_new_offset]]},
                                {'range': f'AU{next_log_row}:AX{next_log_row}',
                                 'values': [[person_1, person_2, selected_day, day_type_code]]}
                            ]

                            adj_ws.batch_update(updates, value_input_option='USER_ENTERED')

                            # clear caches so next load gets fresh data
                            st.session_state.pop(cache_key, None)
                            st.session_state.pop(f"roster_{mmyy}", None)
                            fetch_sheet_data.clear()
                            st.sidebar.success(f"✅ Swapped day {selected_day}: {person_1} → {person_2}")
                            st.rerun()

                    except Exception as e:
                        st.sidebar.error(f"❌ Swap failed: {e}")

        except Exception as e:
            st.sidebar.warning(f"⚠️ Could not load adjustment tool: {e}")

# --------------------------------------------------
# DEV INTERFACE
# --------------------------------------------------

def _rule_to_sentence(rule):
    if not rule:
        return ""
    cls     = rule.get("class","")
    soft    = rule.get("soft", False)
    penalty = rule.get("penalty", 0)
    prefix  = "🔴 Soft —" if soft else "🔵 Hard —"

    if cls == "value":
        s1  = rule.get("subject1","person")
        op  = rule.get("operator","=")
        n   = rule.get("number","")
        s2  = rule.get("subject2","D")
        per = rule.get("per","month")
        op_word = {"=":"exactly","<=":"at most",">=":"at least"}.get(op, op)
        penalty_str = f" *(penalty: {penalty})*" if soft else ""
        return f"{prefix} Each **{s1}** must have {op_word} **{n}** **{s2}** per **{per}**{penalty_str}"

    elif cls == "allow":
        cdt = rule.get("condition_day_type","weekend")
        lg  = rule.get("logic","cannot")
        adt = rule.get("action_day_type","weekend")
        return f"{prefix} If person worked a **{cdt}** last month, they **{lg}** work a **{adt}** this month"

    elif cls == "gap":
        ft   = rule.get("from_type","D")
        tt   = rule.get("to_type","D")
        days = rule.get("days","")
        return f"{prefix} Between **{ft}** and **{tt}** must be at least **{days}** days"

    elif cls == "grouping":
        trait = rule.get("trait","")
        logic = rule.get("logic","must")
        penalty_str = f" *(penalty: {penalty})*" if soft else ""
        trait_map = {
            "same_gender":  "Same gender",
            "partners":     "Partners",
            "same_branch":  "Same branch",
            "drivers":      "Drivers",
        }
        # built-in traits use the map; custom traits display as-is with quotes
        trait_str = trait_map.get(trait, f'"{trait}"')
        if logic == "must_match_d":
            return f"{prefix} **{trait_str}** of S must match D on the same day"
        return f"{prefix} **{trait_str}** **{logic}** be together{penalty_str}"

    return ""

if role == 'Dev':
    st.title("🔧 Dev Panel")
    st.sidebar.button("Logout", on_click=logout, key="dev_logout")

    client = get_gspread_auth()

    # ── Password Management ──
    st.subheader("🔑 Password Management")
    with st.container(border=True):
        try:
            _dev_cfg = fetch_config(client, "MASTER SHEET")
            _cur_admin = _dev_cfg.get("_passwords", {}).get("admin_password", "")
            _cur_user  = _dev_cfg.get("_passwords", {}).get("user_password", "")
        except:
            _cur_admin = ""
            _cur_user  = ""
        new_admin_pw = st.text_input("New Admin Password", value=_cur_admin, type="password", key="new_admin_pw")
        new_user_pw  = st.text_input("New User Password",  value=_cur_user,  type="password", key="new_user_pw")
        if st.button("💾 Save Passwords", use_container_width=True, key="dev_save_pw"):
            try:
                _dev_sh   = client.open("MASTER SHEET")
                _dev_ws   = _dev_sh.worksheet("CONFIG")
                _dev_rows = _dev_ws.get_all_values()
                _pw_upd   = []
                for i, row in enumerate(_dev_rows):
                    if row and row[0].strip() == "admin_password":
                        _pw_upd.append({"range": f"B{i+1}", "values": [[new_admin_pw]]})
                    if row and row[0].strip() == "user_password":
                        _pw_upd.append({"range": f"B{i+1}", "values": [[new_user_pw]]})
                if _pw_upd:
                    _dev_ws.batch_update(_pw_upd)
                    fetch_config.clear()
                    st.success("✅ Passwords updated!")
                else:
                    st.warning("⚠️ Password rows not found in CONFIG sheet.")
            except Exception as e:
                st.error(f"❌ Failed to save passwords: {e}")

    # ── Trait Groups ──
    st.markdown("---")
    st.subheader("🏷️ Trait Groups")
    with st.container(border=True):
        st.caption("Trait groups appear as options in the user form and in the grouping constraint builder. "
                   "Each person can be assigned to one trait group via the user form.")
        _cur_traits = fetch_trait_options(client, "MASTER SHEET")
        _trait_display = ", ".join(_cur_traits) if _cur_traits else "*(none yet)*"
        st.markdown(f"**Current traits:** {_trait_display}")

        tr_col1, tr_col2 = st.columns(2)
        with tr_col1:
            new_trait_name = st.text_input("Add trait", placeholder="e.g. Alpha", key="dev_new_trait")
            if st.button("➕ Add Trait", use_container_width=True, key="dev_add_trait"):
                if not new_trait_name.strip():
                    st.error("❌ Trait name cannot be empty.")
                elif new_trait_name.strip() in _cur_traits:
                    st.warning(f"⚠️ '{new_trait_name.strip()}' already exists.")
                else:
                    try:
                        _updated = _cur_traits + [new_trait_name.strip()]
                        _tr_ws = client.open("MASTER SHEET").worksheet("CONFIG")
                        _tr_rows = _tr_ws.get_all_values()
                        _trait_row_idx = next((i for i, r in enumerate(_tr_rows)
                                               if r and r[0].strip() == "_TRAITS"), None)
                        if _trait_row_idx is not None:
                            _tr_ws.update_acell(f"B{_trait_row_idx+1}", ", ".join(_updated))
                        else:
                            # append a new _TRAITS row before the KEY section
                            _key_idx = next((i for i, r in enumerate(_tr_rows)
                                             if r and r[0].strip().upper() == "KEY"), len(_tr_rows))
                            _tr_ws.insert_row(["_TRAITS", ", ".join(_updated)], _key_idx + 1)
                        fetch_trait_options.clear()
                        fetch_config.clear()
                        st.success(f"✅ Added trait '{new_trait_name.strip()}'")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to add trait: {e}")

        with tr_col2:
            if _cur_traits:
                del_trait = st.selectbox("Remove trait", options=_cur_traits, key="dev_del_trait")
                if st.button("🗑️ Remove Trait", use_container_width=True, key="dev_remove_trait"):
                    try:
                        _updated = [t for t in _cur_traits if t != del_trait]
                        _tr_ws = client.open("MASTER SHEET").worksheet("CONFIG")
                        _tr_rows = _tr_ws.get_all_values()
                        _trait_row_idx = next((i for i, r in enumerate(_tr_rows)
                                               if r and r[0].strip() == "_TRAITS"), None)
                        if _trait_row_idx is not None:
                            _tr_ws.update_acell(f"B{_trait_row_idx+1}", ", ".join(_updated))
                        fetch_trait_options.clear()
                        fetch_config.clear()
                        st.success(f"✅ Removed trait '{del_trait}'")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Failed to remove trait: {e}")
            else:
                st.caption("No traits to remove yet.")

    # ── Constraint Settings ──
    st.markdown("---")
    st.subheader("⚙️ Constraint Settings")
    try:
        _dev_sheet_cfg = fetch_config(client, "MASTER SHEET")
        if "_error" in _dev_sheet_cfg:
            st.warning(f"⚠️ Could not load CONFIG sheet: {_dev_sheet_cfg['_error']}")
        else:
            _dev_constraint_ids = [k for k in _dev_sheet_cfg.keys() if not k.startswith("_")]
            _dev_hard = {k: v for k, v in _dev_sheet_cfg.items() if not k.startswith("_") and v.get("type","").lower() == "hard"}
            _dev_soft = {k: v for k, v in _dev_sheet_cfg.items() if not k.startswith("_") and v.get("type","").lower() == "soft"}
            _dev_drafts = {}

            with st.expander("🪨 Hard Constraints", expanded=False):
                for cid, cv in _dev_hard.items():
                    cur = cv.get("draft_active", cv.get("active", True))
                    nv = st.toggle(cv.get('label', cid), value=cur, key=f"dev_tog_{cid}")
                    if nv != cur:
                        _dev_drafts[cid] = nv
                    _sentence = _rule_to_sentence(cv.get("rule", {}))
                    if _sentence:
                        st.markdown(_sentence, unsafe_allow_html=True)
                    if cv.get("description",""):
                        st.caption(cv.get("description",""))
                    st.markdown("<hr style='margin: 2px 0 8px 0; border: none; border-top: 1px solid rgba(255,255,255,0.1);'>", unsafe_allow_html=True)

            with st.expander("🪶 Soft Constraints", expanded=False):
                for cid, cv in _dev_soft.items():
                    cur = cv.get("draft_active", cv.get("active", True))
                    nv = st.toggle(cv.get('label', cid), value=cur, key=f"dev_tog_{cid}")
                    if nv != cur:
                        _dev_drafts[cid] = nv
                    _sentence = _rule_to_sentence(cv.get("rule", {}))
                    if _sentence:
                        st.markdown(_sentence, unsafe_allow_html=True)
                    if cv.get("description",""):
                        st.caption(cv.get("description",""))
                    st.markdown("<hr style='margin: 2px 0 8px 0; border: none; border-top: 1px solid rgba(255,255,255,0.1);'>", unsafe_allow_html=True)

            if _dev_drafts:
                try:
                    _dws = client.open("MASTER SHEET").worksheet("CONFIG")
                    _drows = _dws.get_all_values()
                    _dupd = []
                    for i, row in enumerate(_drows):
                        if row and row[0].strip() in _dev_drafts:
                            _dupd.append({"range": f"E{i+1}", "values": [["TRUE" if _dev_drafts[row[0].strip()] else "FALSE"]]})
                    if _dupd:
                        _dws.batch_update(_dupd)
                        fetch_config.clear()
                except Exception as e:
                    st.warning(f"⚠️ Could not save draft: {e}")

            if st.button("✅ Publish Changes", use_container_width=True, key="dev_publish"):
                try:
                    _dws = client.open("MASTER SHEET").worksheet("CONFIG")
                    _drows = _dws.get_all_values()
                    _dpub = []
                    for i, row in enumerate(_drows):
                        if row and row[0].strip() in _dev_constraint_ids:
                            draft_val = row[4].strip() if len(row) > 4 else "TRUE"
                            _dpub.append({"range": f"D{i+1}", "values": [[draft_val]]})
                    if _dpub:
                        _dws.batch_update(_dpub)
                        fetch_config.clear()
                        st.success("✅ Constraints published!")
                except Exception as e:
                    st.error(f"❌ Publish failed: {e}")

            # ── Add New Constraint (Rule Builder) ──
            st.markdown("---")
            st.subheader("➕ Add New Constraint")
            with st.container(border=True):
                import json as _json

                _DAY_TYPES = ["weekday", "friday", "weekend", "holiday"]
                _TRAITS    = ["same_gender", "partners", "same_branch", "drivers"]
                _TRAIT_LBL = {"same_gender":"Same Gender","partners":"Partners","same_branch":"Same Branch","drivers":"Drivers"}
                _OPERATORS = ["=", "<=", ">="]
                _SUBJECTS1 = ["person", "day"]
                _SUBJECTS2 = ["D", "S"]
                _LOGICS_ALLOW    = ["can", "cannot"]
                _LOGICS_GROUPING = ["must", "cannot", "must_match_d"]
                _PER_OPTIONS     = ["day", "week", "month"]
                _CLASSES  = ["value", "allow", "gap", "grouping"]
                _CLASS_LBL = {"value":"Value (>, < or =)","allow":"Allow (can/cannot based on condition)","gap":"Gap (minimum days between)","grouping":"Grouping (pairs/traits)"}

                nc_top1, nc_top2, nc_top3 = st.columns(3)
                with nc_top1:
                    nc_type       = st.selectbox("Constraint Type", ["hard","soft"], key="nc_type")
                    nc_duty_type  = st.selectbox("Assignment", ["D","S","DS"], key="nc_duty_type")
                with nc_top2:
                    nc_cls        = st.selectbox("Class", _CLASSES, format_func=lambda x: _CLASS_LBL[x], key="nc_cls")
                with nc_top3:
                    nc_label      = st.text_input("Label", key="nc_label", placeholder="e.g. Max 1 duty per week")
                    nc_desc       = st.text_input("Description", key="nc_desc", placeholder="What does this constraint do?")

                st.markdown("---")
                _preview_rule = {"class": nc_cls, "soft": nc_type=="soft"}

                # ── Class-specific fields ──
                if nc_cls == "value":
                    vc1, vc2, vc3, vc4, vc5 = st.columns(5)
                    with vc1: nc_subj1 = st.selectbox("Subject 1", _SUBJECTS1, key="nc_v_subj1")
                    with vc2: nc_op    = st.selectbox("Operator", _OPERATORS, key="nc_v_op")
                    with vc3: nc_num   = st.number_input("Number", min_value=0, value=1, key="nc_v_num")
                    with vc4: nc_subj2 = st.selectbox("Subject 2", _SUBJECTS2, key="nc_v_subj2",
                                                       index=0 if nc_duty_type=="D" else 1)
                    with vc5: nc_per   = st.selectbox("Per", _PER_OPTIONS, key="nc_v_per")
                    if nc_type == "soft":
                        nc_penalty = st.number_input("Penalty", min_value=0, value=100, step=10, key="nc_v_penalty")
                    else:
                        nc_penalty = 0
                    _preview_rule.update({"subject1":nc_subj1,"operator":nc_op,"number":nc_num,
                                          "subject2":nc_subj2,"per":nc_per,"penalty":nc_penalty})

                elif nc_cls == "allow":
                    ac1, ac2, ac3 = st.columns(3)
                    with ac1: nc_cond_dt   = st.selectbox("Condition Day Type", _DAY_TYPES, key="nc_a_cond")
                    with ac2: nc_logic_a   = st.selectbox("Logic", _LOGICS_ALLOW, key="nc_a_logic")
                    with ac3: nc_action_dt = st.selectbox("Action Day Type", _DAY_TYPES, key="nc_a_action")
                    nc_penalty = 0
                    _preview_rule.update({"condition_day_type":nc_cond_dt,"logic":nc_logic_a,
                                          "action_day_type":nc_action_dt,"penalty":0})

                elif nc_cls == "gap":
                    gc1, gc2, gc3 = st.columns(3)
                    with gc1: nc_from = st.selectbox("From", ["D","S"], key="nc_g_from")
                    with gc2: nc_to   = st.selectbox("To",   ["D","S"], key="nc_g_to")
                    with gc3: nc_days = st.number_input("Days", min_value=1, value=2, key="nc_g_days")
                    nc_penalty = 0
                    _preview_rule.update({"from_type":nc_from,"to_type":nc_to,"days":nc_days,"penalty":0})

                elif nc_cls == "grouping":
                    _BUILTIN_TRAITS    = ["same_gender", "partners", "same_branch", "drivers"]
                    _BUILTIN_TRAIT_LBL = {"same_gender":"Same Gender","partners":"Partners","same_branch":"Same Branch","drivers":"Drivers"}
                    _live_traits       = fetch_trait_options(client, "MASTER SHEET")
                    _all_traits        = _BUILTIN_TRAITS + _live_traits
                    def _trait_fmt(x):
                        return _BUILTIN_TRAIT_LBL.get(x, f'Trait group: "{x}"')
                    grc1, grc2 = st.columns(2)
                    with grc1: nc_trait   = st.selectbox("Trait", _all_traits,
                                                          format_func=_trait_fmt, key="nc_gr_trait")
                    with grc2: nc_logic_g = st.selectbox("Logic", _LOGICS_GROUPING, key="nc_gr_logic")
                    if nc_type == "soft":
                        nc_penalty = st.number_input("Penalty", min_value=0, value=100, step=10, key="nc_gr_penalty")
                    else:
                        nc_penalty = 0
                    _preview_rule.update({"trait":nc_trait,"logic":nc_logic_g,"penalty":nc_penalty})

                nc_param_label = st.text_input("Param Label (optional)", key="nc_param_label",
                                               placeholder="e.g. Gap days")

                # preview
                _prev_str = _rule_to_sentence(_preview_rule)
                if _prev_str:
                    st.markdown("**Preview:** " + _prev_str, unsafe_allow_html=True)

                if st.button("➕ Add Constraint", use_container_width=True, key="nc_add"):
                    if not nc_label:
                        st.error("❌ Label is required.")
                    else:
                        try:
                            _dws = client.open("MASTER SHEET").worksheet("CONFIG")
                            _drows = _dws.get_all_values()
                            existing_ids = [r[0].strip() for r in _drows
                                            if r and r[0].strip()
                                            and r[0].strip().upper() not in ("CONSTRAINT_ID","KEY")]
                            if nc_type == "hard" and nc_duty_type in ("S","DS"):
                                nums = [int(i[2:-1]) for i in existing_ids
                                        if i.upper().startswith("HC") and i.upper().endswith("S")
                                        and i[2:-1].isdigit()]
                                new_cid = f"HC{max(nums)+1 if nums else 1}S"
                            elif nc_type == "hard":
                                nums = [int(i[2:]) for i in existing_ids
                                        if i.upper().startswith("HC") and not i.upper().endswith("S")
                                        and i[2:].isdigit()]
                                new_cid = f"HC{max(nums)+1 if nums else 1}"
                            else:
                                nums = [int(i[2:]) for i in existing_ids
                                        if i.upper().startswith("SC") and i[2:].isdigit()]
                                new_cid = f"SC{max(nums)+1 if nums else 1}"

                            insert_row = len(_drows) + 1
                            for i, row in enumerate(_drows):
                                if row and row[0].strip().upper() == "KEY":
                                    insert_row = i + 1
                                    break

                            _dws.insert_row([], insert_row)
                            new_row_data = [
                                new_cid, nc_label, nc_type, "TRUE", "TRUE",
                                _json.dumps(_preview_rule),
                                nc_param_label, nc_duty_type, nc_cls, nc_desc
                            ]
                            _dws.update(f"A{insert_row}:J{insert_row}", [new_row_data])
                            fetch_config.clear()
                            st.success(f"✅ Added **{new_cid}**: {nc_label}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to add constraint: {e}")
    except Exception as e:
        st.error(f"❌ Could not load constraints: {e}")

# --------------------------------------------------
# USER INTERFACE
# --------------------------------------------------

if role == 'User':
    st.title("🚀 Duty Planner")

    user_page = st.sidebar.segmented_control(
        "",
        options=["✏️ Planning", "🗓️ Viewer"],
    )

    client = get_gspread_auth()

    if user_page == "✏️ Planning":
    
        # initialize session state for date history if not present
        if 'hist_constraints' not in st.session_state:
            st.session_state.hist_constraints = set()
        if 'hist_preferences' not in st.session_state:
            st.session_state.hist_preferences = set()

        _now = date.today(); _opts = [f"{m:02d}{str(y)[2:]}" for y in [_now.year, _now.year+1] for m in range(1,13)]
        view_mmyy = st.selectbox("Month (MMYY)", options=_opts, key="view_mmyy")
        spreadsheet_name = "MASTER SHEET"

        try:
            personal_drive = get_personal_drive_service()
            folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
            gs_query = (
                f"name = '{spreadsheet_name}' "
                f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
                f"and trashed = false "
                f"and '{folder_id}' in parents"
            )
            results = personal_drive.files().list(q=gs_query, fields="files(id)").execute()
            files = results.get('files', [])
            if files:
                st.success(f"✅ Connected to storage!")
            else:
                st.warning(f"⚠️ Connection error: storage failed!")
        except Exception as e:
            st.error(f"❌ Storage failed: {e}")
        
        names_list = fetch_namelist(client, spreadsheet_name)
        st.subheader("Step 1: Select Your Name")
        selected_name = st.selectbox("", options=[""] + names_list)
        if selected_name:
            st.session_state["user_selected_name"] = selected_name

        defaults = {"partner": "None", "driving": "NON-DRIVER", "traits": "", "constraints": "", "preferences": ""}

        if selected_name:
            if "last_fetched_user" not in st.session_state or st.session_state.last_fetched_user != selected_name:
                with st.spinner(f"📦 Retrieving current records for {selected_name}..."):
                    existing = user_engine.get_user_current_data(client, spreadsheet_name, view_mmyy, selected_name)
                    if existing:
                        st.session_state.user_defaults = existing
                        st.session_state.last_fetched_user = selected_name
                        st.session_state.hist_constraints = set(user_engine.parse_string_to_days(existing.get('constraints', ""), view_mmyy))
                        st.session_state.hist_preferences = set(user_engine.parse_string_to_days(existing.get('preferences', ""), view_mmyy))
                        st.toast(f"Loaded data for {selected_name}")
            
            if "user_defaults" in st.session_state:
                defaults = st.session_state.user_defaults

        # date picker with calendar

        st.subheader("Step 2: Pick Your Dates")
        tab1, tab2 = st.tabs(["❌ Constraints (X)", "✅ Duty Days (D)"])

        with tab1:
        
            c_input = st.date_input("Select Date or Range", value=[], key="c_picker")
            c_col1, c_col2 = st.columns(2)
            
            if c_col1.button("➕ Add Constraint"):
                if isinstance(c_input, (list, tuple)):
                    if len(c_input) == 2: # date range
                        curr = c_input[0]
                        while curr <= c_input[1]:
                            st.session_state.hist_constraints.add(curr)
                            curr += timedelta(days=1)
                    elif len(c_input) == 1: # single date
                        st.session_state.hist_constraints.add(c_input[0])
                st.rerun()

            if c_col2.button("🗑️ Reset to Saved (X)"):
                # resets it back to the original spreadsheet data
                st.session_state.hist_constraints = set(user_engine.parse_string_to_days(defaults['constraints'], view_mmyy))
                st.rerun()
            
            constraints_string = user_engine.format_date_list(st.session_state.hist_constraints)
            st.caption(f"Current: {constraints_string if constraints_string else 'None'}")

        with tab2:
            p_input = st.date_input("Select Date or Range", value=[], key="p_picker")
            p_col1, p_col2 = st.columns(2)
            
            if p_col1.button("➕ Add Preference"):
                if isinstance(p_input, (list, tuple)):
                    if len(p_input) == 2:
                        curr = p_input[0]
                        while curr <= p_input[1]:
                            st.session_state.hist_preferences.add(curr)
                            curr += timedelta(days=1)
                    elif len(p_input) == 1:
                        st.session_state.hist_preferences.add(p_input[0])
                st.rerun()

            if p_col2.button("🗑️ Reset to Saved (D)"):
                st.session_state.hist_preferences = set(user_engine.parse_string_to_days(defaults['preferences'], view_mmyy))
                st.rerun()
                
            preferences_string = user_engine.format_date_list(st.session_state.hist_preferences)
            st.caption(f"Current: {preferences_string if preferences_string else 'None'}")

        # form section

        with st.form("user_submission_form"):
            st.subheader("Step 3: Finalise Details")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                d_options = ["NON-DRIVER", "DRIVER", "RIDER"]
                d_idx = d_options.index(defaults['driving']) if defaults['driving'] in d_options else 0
                driving_status = st.selectbox("Your Driving Status", options=d_options, index=d_idx)
                
            with col2:
                p_options = ["None"] + names_list
                p_idx = p_options.index(defaults['partner']) if defaults['partner'] in p_options else 0
                selected_partner = st.selectbox("Your Preferred Partner", options=p_options, index=p_idx)

            with col3:
                s_options = ["", "EXCUSED", "SBF", "NEW"]
                s_idx = 0
                selected_status = st.multiselect("Your Status (If Applicable)", options=s_options, default=s_options[s_idx])

            with col4:
                excused_reason = st.text_input("Reason (if EXCUSED)", placeholder="e.g. Medical appointment...")
                if excused_reason and "EXCUSED" in selected_status:
                    status_string = ", ".join(selected_status) + f" ({excused_reason})"
                else:
                    status_string = ", ".join(selected_status)

            with col5:
                _trait_opts = fetch_trait_options(client, spreadsheet_name)
                _trait_dropdown = [""] + _trait_opts  # blank = no trait assigned
                _cur_trait = defaults.get("traits", "")
                _t_idx = _trait_dropdown.index(_cur_trait) if _cur_trait in _trait_dropdown else 0
                selected_traits = st.selectbox("Your Trait Group", options=_trait_dropdown,
                                               index=_t_idx,
                                               format_func=lambda x: x if x else "— none —")

            final_constraints = st.text_input("Constraints (X)", value=constraints_string)
            final_preferences = st.text_input("Duty Days (D)", value=preferences_string)

            if st.form_submit_button("Save Changes"):
                if not selected_name:
                    st.error("❌ Please select your name before saving.")
                else:
                    with st.spinner("💾 Writing to Google Sheets..."):
                        success, logs = user_engine.update_user_data(
                            client, spreadsheet_name, view_mmyy, 
                            selected_name, selected_partner, 
                            driving_status, selected_traits, final_constraints, final_preferences, status_string
                        )
                        if success:
                            st.success("Preferences updated successfully!")
                            # clear session state on success
                            st.session_state.hist_constraints = []
                            st.session_state.hist_preferences = []
                            if "user_defaults" in st.session_state: 
                                del st.session_state.user_defaults
                            st.rerun()
                        else:
                            st.error(f"❌ Failed to update: {logs[0]}")

    if user_page == "🗓️ Viewer":

        _now = date.today(); _opts = [f"{m:02d}{str(y)[2:]}" for y in [_now.year, _now.year+1] for m in range(1,13)]
        mmyy = st.selectbox("Month/Year (MMYY) to view", options=_opts, key="view_mmyy2")
        spreadsheet_name = "MASTER SHEET"

        curr_m, curr_y = int(mmyy[:2]), int(mmyy[2:])

        st.info(f"Viewing **{mmyy}**!")

        try:
            personal_drive = get_personal_drive_service()
            folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
            sheet_id = fetch_spreadsheet_id(personal_drive, folder_id, spreadsheet_name)
            if sheet_id:
                st.success(f"✅ Connected to storage!")
            else:
                st.warning(f"⚠️ Connection error: storage failed!")
        except Exception as e:
            st.error(f"❌ Storage failed: {e}")

        roster_cache_key = f"roster_user_{mmyy}"
        if roster_cache_key not in st.session_state:
            roster_data, sheet_used, err = user_engine.calendar_view(client, spreadsheet_name, mmyy)
            if not err:
                st.session_state[roster_cache_key] = {"roster_data": roster_data, "sheet_used": sheet_used, "err": err}
        else:
            cached_r = st.session_state[roster_cache_key]
            roster_data, sheet_used, err = cached_r["roster_data"], cached_r["sheet_used"], cached_r["err"]

        sh = client.open(spreadsheet_name)
        
        if err:
            st.warning(f"⚠️ Roster not yet finalised or accessible: {err}")
        else:

            if sheet_used == "D":
                st.success("✅ Showing finalised roster")
            elif sheet_used == "C":
                st.info("ℹ️ Showing draft constraints — roster not yet finalised")

            # 1. Date Math & Setup
            first_day = date(curr_y, curr_m, 1)
            start_padding = (first_day.weekday()) % 7 
            num_days = calendar.monthrange(curr_y, curr_m)[1]
            days_of_week = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

            # 2. Hybrid CSS: Pill for Duty, Plain for Standby
            st.markdown("""
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@400;600;700&display=swap');

                    .cal-container {
                        border: 1px solid #444;
                        border-radius: 15px;
                        overflow: hidden; 
                        margin-top: 10px;
                        font-family: 'Source Sans Pro', sans-serif !important;
                    }
                    
                    .cal-table { 
                        width: 100%; 
                        border-collapse: collapse; 
                        table-layout: fixed;
                        background-color: #262730; /* Darker background to make white text/blue pills pop */
                    }

                    .cal-th { 
                        background-color: #000000; 
                        color: #ffffff; 
                        padding: 12px; 
                        text-align: center; 
                        border-bottom: 1px solid #444;
                        font-weight: 600 !important;
                    }

                    .cal-td { 
                        vertical-align: top; 
                        border: 0.5px solid rgba(255, 255, 255, 0.1); 
                        height: 125px; 
                        padding: 10px; 
                    }

                    .day-num { 
                        font-weight: 700 !important; 
                        font-size: 1rem; 
                        margin-bottom: 10px; 
                        display: block;
                        color: #ffffff;
                    }

                    /* 🚨 DUTY: The "Pill" Style */
                    .duty-item { 
                        font-family: 'Source Sans Pro', sans-serif !important;
                        font-size: 10px; 
                        line-height: 1.2; 
                        margin-bottom: 6px; 
                        font-weight: 600 !important;
                        white-space: nowrap;
                        overflow: hidden;
                        text-overflow: ellipsis;
                        display: block;
                        padding: 4px 8px;
                        border-radius: 6px;
                        background-color: #007bff; 
                        border: 1px solid #0056b3;
                        color: white !important;
                    }
                    
                    /* ⏳ STANDBY: Plain White Text Style */
                    .standby-item { 
                        font-family: 'Source Sans Pro', sans-serif !important;
                        font-size: 11px; 
                        line-height: 1.4; 
                        margin-bottom: 3px; 
                        font-weight: 400 !important;
                        white-space: nowrap;
                        overflow: hidden;
                        text-overflow: ellipsis;
                        display: block;
                        color: white !important; /* Plain white text, no background */
                        padding-left: 2px;
                    }
                        
                    .standby-highlight { 
                        font-family: 'Source Sans Pro', sans-serif !important;
                        font-size: 11px; 
                        line-height: 1.4; 
                        margin-bottom: 3px; 
                        font-weight: 700 !important;
                        white-space: nowrap;
                        overflow: hidden;
                        text-overflow: ellipsis;
                        display: block;
                        color: #e67e22 !important;
                        padding-left: 2px;
                    }
                </style>
            """, unsafe_allow_html=True)

            # 3. Build Table
            html_table = '<div class="cal-container"><table class="cal-table"><thead><tr>'
            for day_name in days_of_week:
                html_table += f'<th class="cal-th">{day_name}</th>'
            html_table += '</tr></thead><tbody><tr>'

            for i in range(start_padding):
                html_table += '<td class="cal-td"></td>'

            current_col = start_padding

            _highlight = st.session_state.get("user_selected_name", "")
            
            for day in range(1, num_days + 1):
                if current_col == 7:
                    html_table += '</tr><tr>'
                    current_col = 0
                
                day_info = roster_data.get(str(day), {"duty": [], "standby": []})
                
                cell_content = f'<span class="day-num">{day}</span>'
                
                # Duty names get the blue pill
                for d_name in day_info["duty"]:
                    _duty_style = 'background:#e67e22;' if d_name == _highlight else ''
                    cell_content += f'<div class="duty-item" style="{_duty_style}" title="Duty: {d_name}">{d_name}</div>'
                
                # Standby names get plain white text
                for s_name in day_info["standby"]:
                    _sb_class = 'standby-highlight' if s_name == _highlight else 'standby-item'
                    cell_content += f'<div class="{_sb_class}" title="Standby: {s_name}">{s_name}</div>'
                
                html_table += f'<td class="cal-td">{cell_content}</td>'
                current_col += 1

            while current_col < 7:
                html_table += '<td class="cal-td"></td>'
                current_col += 1

            html_table += '</tr></tbody></table></div>'

            st.markdown(html_table, unsafe_allow_html=True)