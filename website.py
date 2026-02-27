import streamlit as st
import pandas as pd
import gspread
import planner_engine
import user_engine
import traceback
from datetime import date, timedelta
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

st.set_page_config(page_title="Duty Planner", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# --------------------------------------------------
# AUTHENTICATION
# --------------------------------------------------

def get_gspread_auth():
    """Service account authentication for Sheets (gspread client only)."""
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
    """Builds a Drive service authenticated as your personal Google account.
    Requires [personal_account] section in secrets.toml with:
        token, refresh_token, client_id, client_secret, token_uri
    """
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
# CONVERT FILES FROM .XLSX TO SHEETS
# --------------------------------------------------

def convert_if_excel(client, spreadsheet_name):
    """Uses personal Drive account to find/convert files so storage hits personal quota."""
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

ADMIN_PASSWORD = "password"
USER_PASSWORD = "weapons"

# --------------------------------------------------
# LOGIN
# --------------------------------------------------

def logout():
    st.session_state['logged_in'] = False
    st.session_state['user_role'] = 'User'

if not st.session_state['logged_in']:
    st.title("🚀 Duty Planner")
    st.write("Please select your access level to continue.")

    st.subheader("User")
    st.write("Input your constraints")
    user_pwd = st.text_input("Enter Password", type="password", key="user_password")
    if st.button("Login as User", use_container_width=True):
        if user_pwd == USER_PASSWORD:
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'User'
            st.rerun()
        else:
            st.error("❌ Incorrect Password")

    st.subheader("Admin")
    st.write("Plan rostering")
    admin_pwd = st.text_input("Enter Password", type="password", key="admin_password")
    if st.button("Login as Admin", use_container_width=True):
        if admin_pwd == ADMIN_PASSWORD:
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'Admin'
            st.rerun()
        else:
            st.error("❌ Incorrect Password")
    st.stop()

st.sidebar.button("Logout", on_click=logout)
role = st.session_state['user_role']

# --------------------------------------------------
# ADMIN INTERFACE
# --------------------------------------------------

if "hard4_initialised" not in st.session_state:
    st.session_state["hard4_slider"] = 4
    st.session_state["hard4_input"] = 4
    st.session_state["hard4_initialised"] = True

if role == 'Admin':

# planning parameters

    def update_slider(key):
        st.session_state[key + "_slider"] = st.session_state[key + "_input"]

    def update_input(key):
        st.session_state[key + "_input"] = st.session_state[key + "_slider"]

    st.sidebar.title("📅 Planning Settings")

    st.sidebar.subheader("⚖️ Soft Constraint Weights")

    col_slider, col_input = st.sidebar.columns([3, 1])

    if "S1_slider" not in st.session_state:
        st.session_state["S1_slider"] = 100
        st.session_state["S1_input"] = 100
    if "S2_slider" not in st.session_state:
        st.session_state["S2_slider"] = 60
        st.session_state["S2_input"] = 60
    if "S3_slider" not in st.session_state:
        st.session_state["S3_slider"] = 10
        st.session_state["S3_input"] = 10
    if "S4_slider" not in st.session_state:
        st.session_state["S4_slider"] = 400
        st.session_state["S4_input"] = 400

    with col_slider:
        s1_val = st.slider("Follow Pairings", 0, 500, key="S1_slider", on_change=update_input, args=("S1",), step=10)
        s2_val = st.slider("Different Branches", 0, 500, key="S2_slider", on_change=update_input, args=("S2",), step=10)
        s3_val = st.slider("Driver Mix", 0, 500, key="S3_slider", on_change=update_input, args=("S3",), step=10)
        s4_val = st.slider("Minimum 1x D", 0, 500, key="S4_slider", on_change=update_input, args=("S4",), step=10)

    with col_input:
        s1_manual = st.number_input("Label", 0, 500, key="S1_input", on_change=update_slider, args=("S1",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        s2_manual = st.number_input("Label", 0, 500, key="S2_input", on_change=update_slider, args=("S2",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        s3_manual = st.number_input("Label", 0, 500, key="S3_input", on_change=update_slider, args=("S3",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        s4_manual = st.number_input("Label", 0, 500, key="S4_input", on_change=update_slider, args=("S4",), label_visibility="collapsed")

    config = {
        "S1": s1_val,
        "S2": s2_val,
        "S3": s3_val,
        "S4": s4_val
    }

    st.sidebar.markdown("---")
    st.sidebar.subheader("💯 Point Allocations")

    col_slider, col_input = st.sidebar.columns([3, 1])

    if "weekday_slider" not in st.session_state:
        st.session_state["weekday_slider"] = 1.0
        st.session_state["weekday_input"] = 1.0
    if "friday_slider" not in st.session_state:
        st.session_state["friday_slider"] = 1.5
        st.session_state["friday_input"] = 1.5
    if "weekend_slider" not in st.session_state:
        st.session_state["weekend_slider"] = 2.0
        st.session_state["weekend_input"] = 2.0
    if "holiday_slider" not in st.session_state:
        st.session_state["holiday_slider"] = 2.0
        st.session_state["holiday_input"] = 2.0

    with col_slider:
        weekday_val = st.slider("Weekday Points", 0.0, 10.0, key="weekday_slider", on_change=update_input, args=("weekday",), step=0.5)
        friday_val = st.slider("Friday Points", 0.0, 10.0, key="friday_slider", on_change=update_input, args=("friday",), step=0.5)
        weekend_val = st.slider("Weekend Points", 0.0, 10.0, key="weekend_slider", on_change=update_input, args=("weekend",), step=0.5)
        holiday_val = st.slider("Holiday Points", 0.0, 10.0, key="holiday_slider", on_change=update_input, args=("holiday",), step=0.5)

    with col_input:
        weekday_manual = st.number_input("Label", 0.0, 10.0, key="weekday_input", on_change=update_slider, args=("weekday",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        friday_manual = st.number_input("Label", 0.0, 10.0, key="friday_input", on_change=update_slider, args=("friday",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        weekend_manual = st.number_input("Label", 0.0, 10.0, key="weekend_input", on_change=update_slider, args=("weekend",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        holiday_manual = st.number_input("Label", 0.0, 10.0, key="holiday_input", on_change=update_slider, args=("holiday",), label_visibility="collapsed")

    point_allocations = {
        "weekday_points": weekday_val,
        "friday_points": friday_val,
        "weekend_points": weekend_val,
        "holiday_points": holiday_val
    }

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔓 Model Constraints")

    col_slider, col_input = st.sidebar.columns([3, 1])

    if "hard1_slider" not in st.session_state:
        st.session_state["hard1_slider"] = 2
        st.session_state["hard1_input"] = 2
    if "hard5_slider" not in st.session_state:
        st.session_state["hard5_slider"] = 3
        st.session_state["hard5_input"] = 3
    if "hard1s_slider" not in st.session_state:
        st.session_state["hard1s_slider"] = 2
        st.session_state["hard1s_input"] = 2
    if "hard2s_slider" not in st.session_state:
        st.session_state["hard2s_slider"] = 2
        st.session_state["hard2s_input"] = 2
    if "scalefactor_slider" not in st.session_state:
        st.session_state["scalefactor_slider"] = 4
        st.session_state["scalefactor_input"] = 4

    with col_slider:
        hard1_val = st.slider("Number of Duties Per Day", 0, 4, key="hard1_slider", on_change=update_input, args=("hard1",), step=1)
        hard4_val = st.slider("Gap Between Duties", 0, 10, key="hard4_slider", on_change=update_input, args=("hard4",), step=1)
        hard5_val = st.slider("Maximum No. of Duties", 0, 5, key="hard5_slider", on_change=update_input, args=("hard5",), step=1)
        hard1s_val = st.slider("Number of Standbys Per Day", 0, 4, key="hard1s_slider", on_change=update_input, args=("hard1s",), step=1)
        hard2s_val = st.slider("Gap between S and/or D", 0, 10, key="hard2s_slider", on_change=update_input, args=("hard2s",), step=1)
        scalefactor_val = st.slider("Normalisation Scale", 0, 5, key="scalefactor_slider", on_change=update_input, args=("scalefactor",), step=1)

    with col_input:
        hard1_manual = st.number_input("Label", 0, 4, key="hard1_input", on_change=update_slider, args=("hard1",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        hard4_manual = st.number_input("Label", 0, 10, key="hard4_input", on_change=update_slider, args=("hard4",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        hard5_manual = st.number_input("Label", 0, 5, key="hard5_input", on_change=update_slider, args=("hard5",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        hard1s_manual = st.number_input("Label", 0, 4, key="hard1s_input", on_change=update_slider, args=("hard1s",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        hard2s_manual = st.number_input("Label", 0, 10, key="hard2s_input", on_change=update_slider, args=("hard2s",), label_visibility="collapsed")
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        scalefactor_manual = st.number_input("Label", 0, 5, key="scalefactor_input", on_change=update_slider, args=("scalefactor",), label_visibility="collapsed")

    model_constraints = {
        "hard1": hard1_val,
        "hard4": hard4_val,
        "hard5": hard5_val,
        "hard1s": hard1s_val,
        "hard2s": hard2s_val,
        "scalefactor": scalefactor_val
    }

# main interface

    st.title("🚀 Duty Planner")

    mmyy = st.text_input("Month/Year (MMYY)", value="0126")
    spreadsheet_name = f"Plan_Duty_{mmyy}"

    try:
        client = get_gspread_auth()
        st.success("✅ Connected to Google Account")
    except Exception as e:
        st.error(f"❌ Connection Error: {e}")
        st.stop()

    curr_m, curr_y = int(mmyy[:2]), int(mmyy[2:])
    m_new = curr_m + 1 
    if m_new > 12:
        m_new = 1
        y_new = curr_y + 1
    else:
        y_new = curr_y

    next_file_display = f"{m_new:02d}{y_new:02d}C"

    st.info(f"Targeting Spreadsheet: **{spreadsheet_name}** | Sheet: **{mmyy}C** | New Month File: **{next_file_display}**")

    if st.button("🔄 Convert / Load Spreadsheet"):
        try:
            sh = convert_if_excel(client, spreadsheet_name)
            st.session_state['sh'] = sh
            st.success(f"✅ Loaded: {sh.title}")
        except Exception:
            st.error("🚨 Failed to load spreadsheet")
            st.code(traceback.format_exc())

    if st.button("🔥 Run Optimiser"):
        try:

            sh = convert_if_excel(client, spreadsheet_name)

            with st.spinner("📥 Fetching Sheet Data..."):

                def get_df(sheet_name, header_row, use_cols = None):
                    try:
                        data = sh.worksheet(sheet_name).get_all_values()
                        df = pd.DataFrame(data)
                        df.columns = df.iloc[0]
                        df = df[1:].reset_index(drop=True)
                        if use_cols:
                            df = df.iloc[:, :use_cols]
                        return df.head(250)
                    except Exception as e:
                        raise ValueError(f"Error loading sheet '{sheet_name}': {e}")

                constraints_raw = get_df(f"{mmyy}C", header_row=1)
                constraints_raw.iloc[:, 43] = pd.to_numeric(constraints_raw.iloc[:, 43], errors='coerce').fillna(0)
                holidays_raw = get_df("Holiday", header_row=0, use_cols=3)
                partners_raw = get_df("Partners", header_row=0, use_cols=5)
                namelist_raw = get_df("Namelist", header_row=0, use_cols=3)

                try:
                    prev_m = curr_m - 1 if curr_m > 1 else 12
                    prev_y = curr_y if curr_m > 1 else curr_y - 1
                    last_month_raw = get_df(f"{prev_m:02d}{prev_y:02d}D", header_row=1)
                except:
                    st.warning("Previous month data not found.")
                    last_month_raw = None
            

            with st.spinner("🧠 Solving Optimisation..."):

                data_bundle = {
                    "constraints": constraints_raw,
                    "holidays": holidays_raw,
                    "year": 2000 + int(mmyy[2:]),
                    "year_old": 2000 + prev_y,
                    "month": int(mmyy[:2]),
                    "month_old": prev_m,
                    "partners": partners_raw,
                    "namelist": namelist_raw,
                    "last_month": last_month_raw
                }

                planned_df, n_scale, ranges = planner_engine.run_optimisation(data_bundle, config, point_allocations, model_constraints)

                if planned_df is not None:
                    st.session_state['planned_df'] = planned_df
                    st.session_state['n_scale'] = n_scale
                    st.session_state['ranges'] = ranges
                    st.session_state['active_sh_name'] = sh.title

                    st.success("✅ Optimisation Successful!")
                    state = "complete"
                else:
                    st.warning("❌ No Solution Found")
                    state = "error"

        except Exception:
            st.error("🚨 Critical Error Detected")
            st.code(traceback.format_exc())


    # planning buttons

    final_name = st.session_state.get('active_sh_name', spreadsheet_name)

    if st.button("💾 Save to Google Sheets (Output D)"):
        if 'planned_df' in st.session_state:
            # 1. archive the original MMYYC using personal Drive account
            with st.spinner("🧳 Creating Archive..."):
                personal_drive = get_personal_drive_service()
                folder_id = st.secrets["app_config"]["personal_drive_folder_id"]
                planner_engine.archive_source_sheet(client, final_name, mmyy, folder_id, personal_drive)
            
            # 2. write output
            with st.spinner("✏️ Writing Output..."):
                out_name = planner_engine.create_backup_and_output(
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

                sh = client.open(final_name)
                sh.update_title(next_file_name)

                st.success(f"Done! File renamed to **{next_file_name}**")
                st.session_state.pop('planned_df', None)
        else:
            st.warning("Run the optimiser first!")

    st.markdown("---")

    # manual adjustments writing

    st.markdown("### 🔄 Manual Duty Adjustments")

    target_sheet_name = f"{mmyy}C"

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
        if not files:
            raise FileNotFoundError(f"Could not find '{spreadsheet_name}' in your Drive folder")
        sh_admin = client.open_by_key(files[0]['id'])
        names_for_adj = user_engine.get_namelist(client, spreadsheet_name)
        adj_ws = sh_admin.worksheet(target_sheet_name)
        
        with st.container(border=True):
            col_p1, col_p2, col_type = st.columns(3)
            with col_p1:
                person_minus = st.selectbox("Person giving up duty (MINUS)", options=[""] + names_for_adj)
            with col_p2:
                person_plus = st.selectbox("Person taking over duty (ADD)", options=[""] + names_for_adj)
            with col_type:
                day_type = st.selectbox("Day Type", options=["Weekday (WD)", "Friday (F)", "Weekend (WE)", "Holiday (H)"])

            if st.button("+ Record Adjustment", use_container_width=True):
                if not person_minus or not person_plus:
                    st.error("❌ Please select both individuals.")
                elif person_minus == person_plus:
                    st.error("❌ Cannot swap between the same person.")
                else:
                    type_map = {"Weekday (WD)": "WD", "Friday (F)": "F", "Weekend (WE)": "WE", "Holiday (H)": "H"}
                    suffix = type_map[day_type]
                    
                    # find next row in column AW (49)
                    col_aw_values = adj_ws.col_values(49)
                    next_row = max(37, len(col_aw_values) + 1)

                    updates = [
                        {'range': f'AW{next_row}:AX{next_row}', 'values': [[person_minus.upper(), f"MINUS 1X {suffix}"]]},
                        {'range': f'AW{next_row+1}:AX{next_row+1}', 'values': [[person_plus.upper(), f"ADD 1X {suffix}"]]}
                    ]
                    
                    adj_ws.batch_update(updates, value_input_option='USER_ENTERED')
                    st.success(f"📝 Recorded")
                    st.rerun()

        # review and clear adjustments

        st.write("**Current adjustments (AW:AX):**")
        
        raw_adj = adj_ws.get("AW37:AX100")
        if raw_adj:
            adj_df = pd.DataFrame(raw_adj, columns=["Name", "Adjustment"])
            st.dataframe(adj_df, use_container_width=True, hide_index=True)
            
            if st.button("🗑️ Clear Adjustments (Names/Text Only)", type="secondary"):
                adj_ws.batch_clear(["AW37:AX100"])
                st.toast("🚮 Adjustment names and text cleared")
                st.rerun()
        else:
            st.caption(f"No adjustments recorded in {target_sheet_name}")

    except Exception as e:
        st.warning(f"No adjustments found.")

# --------------------------------------------------
# USER INTERFACE
# --------------------------------------------------

if role == 'User':
    st.title("🚀 Duty Planner")
    
    # initialize session state for date history if not present
    if 'hist_constraints' not in st.session_state:
        st.session_state.hist_constraints = []
    if 'hist_preferences' not in st.session_state:
        st.session_state.hist_preferences = []

    client = get_gspread_auth()
    view_mmyy = st.text_input("Month (MMYY)", value="0126")
    spreadsheet_name = f"Plan_Duty_{view_mmyy}"

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
            st.success(f"✅ Found: {spreadsheet_name}")
        else:
            st.warning(f"⚠️ No spreadsheet found for {spreadsheet_name}")
    except Exception as e:
        st.error(f"❌ Drive check failed: {e}")
    
    names_list = user_engine.get_namelist(client, spreadsheet_name)
    selected_name = st.selectbox("Step 1: Select Your Name to Load Data", options=[""] + names_list)

    defaults = {"partner": "None", "driving": "NON-DRIVER", "constraints": "", "preferences": ""}

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

        if c_col2.button("🗑️ Undo Constraints"):
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

        if p_col2.button("🗑️ Clear/Undo Preferences"):
            st.session_state.hist_preferences = set(user_engine.parse_string_to_days(defaults['preferences'], view_mmyy))
            st.rerun()
            
        preferences_string = user_engine.format_date_list(st.session_state.hist_preferences)
        st.caption(f"Current: {preferences_string if preferences_string else 'None'}")

    # form section

    with st.form("user_submission_form"):
        st.subheader("Step 3: Finalize Details")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            d_options = ["NON-DRIVER", "DRIVER", "RIDER"]
            d_idx = d_options.index(defaults['driving']) if defaults['driving'] in d_options else 0
            driving_status = st.selectbox("Your Driving Status", options=d_options, index=d_idx)
            
        with col2:
            p_options = ["None"] + names_list
            p_idx = p_options.index(defaults['partner']) if defaults['partner'] in p_options else 0
            selected_partner = st.selectbox("Your Preferred Partner", options=p_options, index=p_idx)

        with col3:
            s_options = ["","SBF", "SAIL", "NDP", "EXCUSED", "MEDICAL", "ON COURSE", "NEW"]
            # check if defaults has status_string, otherwise default to empty
            current_status = defaults.get('status_string', "").split(", ") if defaults.get('status_string') else []
            selected_status = st.multiselect("Your Status", options=s_options, default=[s for s in current_status if s in s_options])
            status_string_out = ", ".join(selected_status)
            
        final_constraints = st.text_input("Constraints (X)", value=constraints_string)
        final_preferences = st.text_input("Duty Days (D)", value=preferences_string)

        if st.form_submit_button("Save Changes"):
            with st.spinner("💾 Writing to Google Sheets..."):
                success, logs = user_engine.update_user_data(
                    client, spreadsheet_name, view_mmyy, 
                    selected_name, selected_partner, 
                    driving_status, final_constraints, final_preferences, status_string_out
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