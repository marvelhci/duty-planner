import streamlit as st
import pandas as pd
import gspread
import os
import planner_engine
import user_engine
import traceback
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
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
    """OAuth authentication using only Streamlit secrets."""
    try:
        if "gcp_service_account" in st.secrets:
            # Convert the secrets object to a standard dictionary
            creds_dict = dict(st.secrets["gcp_service_account"])
            
            # Authorize using the service account credentials
            client = gspread.service_account_from_dict(creds_dict)
            
            # Return the client and None for creds (since service accounts handle their own state)
            return client, None
        else:
            st.error("❌ 'gcp_service_account' not found in secrets.toml")
            st.stop()
    except Exception as e:
        st.error(f"❌ Authentication Error: {e}")
        st.stop()

# --------------------------------------------------
# CONVERT FILES FROM .XLSX TO SHEETS
# --------------------------------------------------

def convert_if_excel(client, creds, spreadsheet_name):
    drive_service = build('drive', 'v3', credentials=creds)

    gs_query = f"name = '{spreadsheet_name}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    gs_results = drive_service.files().list(q=gs_query, fields="files(id)").execute()
    gs_files = gs_results.get('files', [])

    if gs_files:
        return client.open_by_key(gs_files[0]['id'])
    
    ex_query = f"(name = '{spreadsheet_name}' or name = '{spreadsheet_name}.xlsx') and mimeType != 'application/vnd.google-apps.spreadsheet' and trashed = false"
    ex_results = drive_service.files().list(q=ex_query, fields="files(id, name)").execute()
    ex_files = ex_results.get('files', [])

    if not ex_files:
        raise FileNotFoundError(f"Could not find any active Excel or Google Sheet named '{spreadsheet_name}'")

    with st.spinner("📦 Excel source detected. Converting to Google Sheets..."):
        target_excel = ex_files[0]
        
        file_metadata = {
            'name': spreadsheet_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }

        converted_file = drive_service.files().copy(
            fileId=target_excel['id'],
            body=file_metadata
        ).execute()

    return client.open_by_key(converted_file.get('id'))

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
            st.error("Incorrect Password")

    st.subheader("Admin")
    st.write("Plan rostering")
    admin_pwd = st.text_input("Enter Password", type="password", key="admin_password")
    if st.button("Login as Admin", use_container_width=True):
        if admin_pwd == ADMIN_PASSWORD:
            st.session_state['logged_in'] = True
            st.session_state['user_role'] = 'Admin'
            st.rerun()
        else:
            st.error("Incorrect Password")
    st.stop()

st.sidebar.button("Logout", on_click=logout)
role = st.session_state['user_role']

# --------------------------------------------------
# ADMIN INTERFACE
# --------------------------------------------------

if role == 'Admin':

# planning parameters

    def update_slider(key):
        st.session_state[key + "_slider"] = st.session_state[key + "_input"]

    def update_input(key):
        st.session_state[key + "_input"] = st.session_state[key + "_slider"]

    st.sidebar.title("📅 Planning Settings")

    mmyy = st.sidebar.text_input("Month/Year (MMYY)", value="0126")
    spreadsheet_name = f"Plan_Duty_{mmyy}"

    st.sidebar.markdown("---")
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
    if "hard4_slider" not in st.session_state:
        st.session_state["hard4_slider"] = 4
        st.session_state["hard4_input"] = 4
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
        "hard1_points": hard1_val,
        "hard4_points": hard4_val,
        "hard5_points": hard5_val,
        "hard1s_points": hard1s_val,
        "hard2s_points": hard2s_val,
        "scalefactor_points": scalefactor_val
    }

# main interface

    st.title("🚀 Duty Planner")

    try:
        client, creds = get_gspread_auth()
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

    if st.button("🔥 Run Optimiser"):
        try:

            sh = convert_if_excel(client, creds, spreadsheet_name)

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

    st.markdown("---")

    final_name = st.session_state.get('active_sh_name', spreadsheet_name)

    if st.button("💾 Save to Google Sheets (Output D)"):
        if 'planned_df' in st.session_state:
            # 1. Archive the original MMYYC
            with st.spinner("🧳 Creating Archive..."):
                planner_engine.archive_source_sheet(client, final_name, mmyy)
            
            # 2. Write output
            with st.spinner("✏️ Writing Output..."):
                out_name = planner_engine.create_backup_and_output(
                    client, final_name, mmyy,
                    st.session_state['planned_df'],
                    st.session_state['n_scale'],
                    st.session_state['ranges']
                )

            # 3. Create Next Month Template
            with st.spinner("⏭️ Preparing Next Month..."):
                _, next_file_name =planner_engine.generate_next_month_template(
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

# --------------------------------------------------
# USER INTERFACE
# --------------------------------------------------

elif role == 'User':
    st.title("🚀 Duty Planner")
    
    client, _ = get_gspread_auth()
    view_mmyy = st.sidebar.text_input("Month (MMYY)", value="0126")
    spreadsheet_name = f"Plan_Duty_{view_mmyy}"
    
    names_list = user_engine.get_namelist(client, spreadsheet_name)
    selected_name = st.selectbox("Step 1: Select Your Name to Load Data", options=[""] + names_list)

    defaults = {"partner": "None", "driving": "NON-DRIVER", "constraints": "", "preferences": ""}

    # fetch data from previous month
    if selected_name:
        if "last_fetched_user" not in st.session_state or st.session_state.last_fetched_user != selected_name:
            with st.spinner(f"📦 Retrieving current records for {selected_name}..."):
                existing = user_engine.get_user_current_data(client, spreadsheet_name, view_mmyy, selected_name)
                if existing:
                    st.session_state.user_defaults = existing
                    st.session_state.last_fetched_user = selected_name
                    st.toast(f"Loaded data for {selected_name}")
        

        if "user_defaults" in st.session_state:
            defaults = st.session_state.user_defaults

    # input form
    with st.form("user_submission_form"):
        st.subheader("Step 2: Update Your Constraints")
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
            s_idx = 0
            selected_status = st.multiselect("Your Status (If Applicable)", options=s_options, default=s_options[s_idx])
            status_string = ", ".join(selected_status)
            
        constraints = st.text_input("Constraints (X) (e.g. 1, 2, 3)", value=defaults['constraints'], help="Comma separated numbers")
        preferences = st.text_input("Duty Days (D) (e.g. 1, 2, 3)", value=defaults['preferences'], help="Comma separated numbers")

        if st.form_submit_button("Save Changes"):
            with st.spinner("💾 Writing to Google Sheets..."):
                success, logs = user_engine.update_user_data(
                    client, spreadsheet_name, view_mmyy, 
                    selected_name, selected_partner, 
                    driving_status, constraints, preferences, status_string
                )
                if success:
                    st.success("Preferences updated successfully!")
                    del st.session_state.user_defaults
                else:
                    st.error(f"Failed to update: {logs[0]}")