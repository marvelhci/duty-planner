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
import requests as http_requests

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
# CACHED DATA FETCHERS
# --------------------------------------------------

@st.cache_data(ttl=300)
def fetch_sheet_data(_client, spreadsheet_name, sheet_name):
    """Cached sheet reader — TTL 5 mins."""
    sh = _client.open(spreadsheet_name)
    return sh.worksheet(sheet_name).get_all_values()

@st.cache_data(ttl=300)
def fetch_namelist(_client, spreadsheet_name):
    """Cached namelist fetch."""
    return user_engine.get_namelist(_client, spreadsheet_name)

@st.cache_data(ttl=600)
def fetch_spreadsheet_id(_personal_drive, folder_id, spreadsheet_name):
    """Cached Drive file ID lookup — TTL 10 mins."""
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
        if "sbf_slider" not in st.session_state:
            st.session_state["sbf_slider"] = 2
            st.session_state["sbf_input"] = 2

        with col_slider:
            hard1_val = st.slider("Number of Duties Per Day", 0, 4, key="hard1_slider", on_change=update_input, args=("hard1",), step=1)
            hard4_val = st.slider("Gap Between Duties", 0, 10, key="hard4_slider", on_change=update_input, args=("hard4",), step=1)
            hard5_val = st.slider("Maximum No. of Duties", 0, 5, key="hard5_slider", on_change=update_input, args=("hard5",), step=1)
            hard1s_val = st.slider("Number of Standbys Per Day", 0, 4, key="hard1s_slider", on_change=update_input, args=("hard1s",), step=1)
            hard2s_val = st.slider("Gap between S and/or D", 0, 10, key="hard2s_slider", on_change=update_input, args=("hard2s",), step=1)
            scalefactor_val = st.slider("Normalisation Scale", 0, 5, key="scalefactor_slider", on_change=update_input, args=("scalefactor",), step=1)
            sbf_val = st.slider("SB Bonus", 0, 5, key="sbf_slider", on_change=update_input, args=("sbf",), step=1)

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
            st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
            sbf_manual = st.number_input("Label", 0, 5, key="sbf_input", on_change=update_slider, args=("sbf",), label_visibility="collapsed")

        model_constraints = {
            "hard1": hard1_val,
            "hard4": hard4_val,
            "hard5": hard5_val,
            "hard1s": hard1s_val,
            "hard2s": hard2s_val,
            "scalefactor": scalefactor_val,
            "sbf_val": sbf_val
        }

        # main interface

        mmyy = st.text_input("Month/Year (MMYY) to plan", value="0126")
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

        next_file_display = f"{m_new:02d}{y_new:02d}"

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
                        namelist_raw = get_df("Namelist", header_row=0, use_cols=4)

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

                        planned_df, n_scale, ranges = planner_engine.run_optimisation(data_bundle, config, point_allocations, model_constraints)

                        if planned_df is not None:
                            st.session_state['planned_df'] = planned_df
                            st.session_state['n_scale'] = n_scale
                            st.session_state['ranges'] = ranges
                            st.session_state['active_sh_name'] = sh.title

                            st.success("✅ Optimisation Successful!")
                            state = "complete"
                        else:
                            st.warning("⚠️ No Solution Found")
                            state = "error"

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

                        st.success(f"✅ Done!")
                        st.session_state.pop('planned_df', None)
                else:
                    st.warning("⚠️ Run the optimiser first!")

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

    if admin_page == "✏️ Editing":

        mmyy = st.text_input("Month/Year (MMYY) to edit", value="0126")
        spreadsheet_name = "MASTER SHEET"

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

        # point allocations — use planning page slider values if available, else defaults
        point_allocations = {
            "weekday_points": st.session_state.get("weekday_slider", 1.0),
            "friday_points": st.session_state.get("friday_slider", 1.5),
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
                        import weasyprint
                        cal_first_day = date(2000 + int(mmyy[2:]), int(mmyy[:2]), 1)
                        cal_num_days = calendar.monthrange(2000 + int(mmyy[2:]), int(mmyy[:2]))[1]
                        cal_start_pad = cal_first_day.weekday()
                        month_label = cal_first_day.strftime("%B %Y")
                        dow = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

                        css_lines = [
                            "body { font-family: Arial, sans-serif; background: white; color: black; margin: 20px; }",
                            "h2 { text-align: center; margin-bottom: 16px; }",
                            "table { width: 100%; border-collapse: collapse; table-layout: fixed; }",
                            "th { background: #333; color: white; padding: 8px; text-align: center; font-size: 12px; }",
                            "td { border: 1px solid #ccc; vertical-align: top; height: 110px; padding: 6px; width: 14.28%; }",
                            ".day-num { font-weight: bold; font-size: 13px; display: block; margin-bottom: 4px; }",
                            ".duty-item { font-size: 9px; background: #1a73e8; color: white; border-radius: 4px; padding: 2px 5px; margin-bottom: 3px; display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; }",
                            ".standby-item { font-size: 10px; color: #333; display: block; margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }",
                        ]
                        css = " ".join(css_lines)

                        cal_html = "<!DOCTYPE html><html><head><meta charset='utf-8'><style>" + css + "</style></head><body>"
                        cal_html += "<h2>" + month_label + " Duty Roster</h2><table><thead><tr>"
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
                            cal_html += '<td><span class="day-num">' + str(day) + "</span>"
                            for n in info["duty"]:
                                cal_html += '<span class="duty-item">' + n + "</span>"
                            for n in info["standby"]:
                                cal_html += '<span class="standby-item">' + n + "</span>"
                            cal_html += "</td>"
                            col_idx += 1

                        while col_idx < 7:
                            cal_html += "<td></td>"
                            col_idx += 1
                        cal_html += "</tr></tbody></table></body></html>"

                        pdf_bytes = weasyprint.HTML(string=cal_html).write_pdf(
                            stylesheets=[weasyprint.CSS(string="@page { size: A4 landscape; margin: 1cm; }")]
                        )
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
                                "https://docs.google.com/spreadsheets/d/" + sid_dl + "/export"
                                "?format=pdf"
                                "&gid=" + str(sheet_gid_dl) +
                                "&portrait=false"
                                "&fitw=true"
                                "&gridlines=true"
                                "&r1=0&c1=0&r2=" + str(last_name_row_dl) + "&c2=45"
                                "&ir=false&ic=false"
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
            st.session_state.hist_constraints = []
        if 'hist_preferences' not in st.session_state:
            st.session_state.hist_preferences = []

        view_mmyy = st.text_input("Month (MMYY)", value="0126")
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
            col1, col2, col3, col4 = st.columns(4)
            
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
                
            final_constraints = st.text_input("Constraints (X)", value=constraints_string)
            final_preferences = st.text_input("Duty Days (D)", value=preferences_string)

            if st.form_submit_button("Save Changes"):
                with st.spinner("💾 Writing to Google Sheets..."):
                    success, logs = user_engine.update_user_data(
                        client, spreadsheet_name, view_mmyy, 
                        selected_name, selected_partner, 
                        driving_status, final_constraints, final_preferences, status_string
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

        mmyy = st.text_input("Month/Year (MMYY) to edit", value="0126")
        spreadsheet_name = "MASTER SHEET"

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