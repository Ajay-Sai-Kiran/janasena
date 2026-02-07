import streamlit as st
import pandas as pd
import requests
import json
import time
from datetime import datetime
import os
import connection
from bs4 import BeautifulSoup
import re

# Set page config
st.set_page_config(
    page_title="Telangana Urban Voter Data Extractor",
    page_icon="🗳️",
    layout="wide"
)

# Initialize Session State
if 'elections' not in st.session_state:
    st.session_state['elections'] = []
if 'districts' not in st.session_state:
    st.session_state['districts'] = []
    
if 'municipalities' not in st.session_state:
    st.session_state['municipalities'] = []
if 'wards' not in st.session_state:
    st.session_state['wards'] = []
    
if 'download_results' not in st.session_state:
    st.session_state['download_results'] = []

BASE_URL = "https://urban2025.tsec.gov.in"

# --- Helper Functions ---
def log_request(url, params=None):
    if 'logs' not in st.session_state:
        st.session_state['logs'] = []
    
    timestamp = datetime.now().strftime("%H:%M:%S")
    msg = f"[{timestamp}] Request: {url}"
    if params:
        msg += f" Params: {params}"
    st.session_state['logs'].insert(0, msg)

def get_session():
    if 'session' not in st.session_state:
        status_text = st.empty()
        status_text.info("Initializing secure session...")
        session = connection.get_session()
        try:
            # Warm up connection to get fresh cookies
            root_url = "https://urban2025.tsec.gov.in/"
            log_request(root_url, "Warming up session...")
            session.get(root_url, timeout=30)
            status_text.success("Session initialized!")
            time.sleep(1)
            status_text.empty()
        except Exception as e:
            st.error(f"Session initialization warning: {e}")
            log_request(root_url, f"Warmup Failed: {e}")
            
        st.session_state['session'] = session
        
    return st.session_state['session']

def fetch_initial_data():
    """Fetch Elections and Districts from the main page HTML"""
    session = get_session()
    url = f"{BASE_URL}/slNoWardWiseVoterlisturbanMapped.do"
    log_request(url)
    try:
        response = session.get(url, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Elections
            elections = []
            # Updated ID: election_id
            election_select = soup.find('select', {'id': 'election_id'})
            if election_select:
                for opt in election_select.find_all('option'):
                    val = opt.get('value')
                    if val and val != '0': # Skip "Select"
                        elections.append({'id': val, 'name': opt.text.strip()})
            
            # Districts
            districts = []
            # Updated ID: district_id
            district_select = soup.find('select', {'id': 'district_id'})
            if district_select:
                for opt in district_select.find_all('option'):
                    val = opt.get('value')
                    if val and val != '0': # Skip "Select"
                        districts.append({'id': val, 'name': opt.text.strip()})
                        
            return elections, districts
            
    except Exception as e:
        st.error(f"Error fetching initial data: {e}")
        log_request(url, f"Error: {e}")
        
        # Fallback Data
        st.warning("⚠️ Connection failed. Using fallback data (Nizamabad / 2026 Election).")
        fallback_elections = [{'id': '186', 'name': 'ORDINARY ELECTIONS TO MUNICIPALITIES AND MUNICIPAL CORPORATIONS, 2026'}]
        fallback_districts = [{'id': '05', 'name': 'Nizamabad'}]
        return fallback_elections, fallback_districts

    return [], []

def fetch_municipalities(district_code):
    session = get_session()
    # Updated Endpoint based on HTML JS: wardwisevoterlisturban.do?mode=getMunicipality&district_id=05
    url = f"{BASE_URL}/wardwisevoterlisturban.do"
    params = {'mode': 'getMunicipality', 'district_id': district_code}
    log_request(url, params)
    try:
        response = session.post(url, params=params, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
            # The response is likely HTML options, not JSON.
            # <option value="1">...</option>
            soup = BeautifulSoup(response.content, 'html.parser')
            options = []
            for opt in soup.find_all('option'):
                val = opt.get('value')
                if val and val != '0':
                    options.append({'id': val, 'name': opt.text.strip()})
            return options
    except Exception as e:
        st.error(f"Error fetching municipalities: {e}")
        log_request(url, f"Error: {e}")
    return []

def fetch_wards(district_code, municipality_code):
    session = get_session()
    # Updated Endpoint based on HTML JS: wardwisevoterlisturban.do?mode=getWard&district_id=..&municipality_id=..
    url = f"{BASE_URL}/wardwisevoterlisturban.do"
    params = {'mode': 'getWard', 'district_id': district_code, 'municipality_id': municipality_code}
    log_request(url, params)
    try:
        response = session.post(url, params=params, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
            # Response is HTML options
            soup = BeautifulSoup(response.content, 'html.parser')
            options = []
            for opt in soup.find_all('option'):
                val = opt.get('value')
                if val and val != '0' and val != '':
                    options.append({'id': val, 'name': opt.text.strip()})
            return options
    except Exception as e:
        st.error(f"Error fetching wards: {e}")
        log_request(url, f"Error: {e}")
    return []

def fetch_ac_parts(ward_code, municipality_code, district_code):
    session = get_session()
    # Updated Endpoint: slNoWardWiseVoterlisturbanMapped.do?mode=getPartNos...
    url = f"{BASE_URL}/slNoWardWiseVoterlisturbanMapped.do"
    params = {
        'mode': 'getPartNos',
        'district_id': district_code,
        'municipality_id': municipality_code,
        'ward_id': ward_code
    }
    log_request(url, params)
    
    try:
        response = session.post(url, params=params, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
             # Response is HTML options
            soup = BeautifulSoup(response.content, 'html.parser')
            parts = []
            for opt in soup.find_all('option'):
                val = opt.get('value')
                if val and val != '0' and val != '':
                    parts.append({'partno': val})
            return parts
    except Exception as e:
        st.error(f"Error fetching AC parts: {e}")
        log_request(url, f"Error: {e}")
    return []

# --- UI ---
st.title("Telangana Urban Voter Data Extractor")
st.markdown("Use the filters below to select the area and generate an Excel report of available voter lists.")

# Load Initial Data
if not st.session_state['elections'] or not st.session_state['districts']:
    with st.spinner("Connecting to TSEC server..."):
        elections, districts = fetch_initial_data()
        st.session_state['elections'] = elections
        st.session_state['districts'] = districts
        if not elections:
             st.warning("Could not fetch elections automatically. Please check your connection.")

col1, col2 = st.columns(2)

with col1:
    # 1. Election
    election_options = {e['id']: e['name'] for e in st.session_state['elections']}
    # default to index 0 if available
    idx = 0
    selected_election_code = st.selectbox(
        "Election", 
        options=list(election_options.keys()), 
        format_func=lambda x: election_options.get(x, x),
        index=idx
    )

    # 2. District
    district_options = {d['id']: d['name'] for d in st.session_state['districts']}
    selected_district_code = st.selectbox(
        "District", 
        options=list(district_options.keys()), 
        format_func=lambda x: district_options.get(x, x)
    )

    # 3. Municipality
    # Retrieve Munis based on district selection
    if 'last_district' not in st.session_state:
        st.session_state['last_district'] = None
        
    if selected_district_code != st.session_state['last_district']:
         # Auto-fetch municipalities when district changes
         if selected_district_code:
            with st.spinner("Fetching Municipalities..."):
                munis = fetch_municipalities(selected_district_code)
                st.session_state['municipalities'] = munis
                st.session_state['last_district'] = selected_district_code
                
    muni_options = {m['id']: m['name'] for m in st.session_state['municipalities']}
    
    selected_muni_code = st.selectbox(
        "Municipality", 
        options=list(muni_options.keys()), 
        format_func=lambda x: muni_options.get(x, x)
    )

with col2:
    # 4. Wards
    # Retrieve Wards based on muni selection
    if 'last_muni' not in st.session_state:
        st.session_state['last_muni'] = None
        
    if selected_muni_code != st.session_state['last_muni']:
         if selected_muni_code:
            with st.spinner("Fetching Wards..."):
                wards = fetch_wards(selected_district_code, selected_muni_code)
                st.session_state['wards'] = wards
                st.session_state['last_muni'] = selected_muni_code
             
    ward_options = {w['id']: w['name'] for w in st.session_state['wards']}
    
    # Selection Mode
    selection_mode = st.radio("Scope", ["Specific Ward", "All Wards in Municipality"])
    
    selected_wards_data = [] 
    
    if selection_mode == "Specific Ward":
        specific_ward_code = st.selectbox(
            "Select Ward", 
            options=list(ward_options.keys()), 
            format_func=lambda x: ward_options.get(x, x)
        )
        if specific_ward_code:
             selected_wards_data = [{'code': specific_ward_code, 'name': ward_options[specific_ward_code]}]
    else:
        # All Wards
        if ward_options:
             selected_wards_data = [{'code': k, 'name': v} for k,v in ward_options.items()]
             st.info(f"Will process {len(selected_wards_data)} wards")
        else:
            st.info("No wards available.")

st.markdown("---")

# Generate Report Button
if st.button("Generate Excel Report", type="primary"):
    if not selected_wards_data:
        st.error("Please select wards to process.")
    else:
        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total = len(selected_wards_data)
        
        for i, ward in enumerate(selected_wards_data):
            status_text.text(f"Processing {ward['name']} ({i+1}/{total})...")
            
            # Fetch parts for this ward
            parts = fetch_ac_parts(ward['code'], selected_muni_code, selected_district_code)
            
            if not parts:
                results.append({
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Election": election_options.get(selected_election_code, selected_election_code),
                    "District": district_options.get(selected_district_code, selected_district_code),
                    "Municipality": muni_options.get(selected_muni_code, selected_muni_code),
                    "Ward Name": ward['name'],
                    "Ward Code": ward['code'],
                    "AC Part No": "N/A",
                    "Status": "No Data Found",
                    "Filename": "-"
                })
            else:
                for part in parts:
                    p_no = str(part['partno'])
                    # Generate Direct Download Link using CORRECT endpoint
                    # URL: slNoWardWiseVoterlisturbanMapped.do?mode=createViewInEnglishReport&election_id=186&district_id=05&mnc_id=1&ward_id=1&circle_id=0&part_no=22
                    download_url = (
                        f"{BASE_URL}/slNoWardWiseVoterlisturbanMapped.do?"
                        f"mode=createViewInEnglishReport&"
                        f"election_id={selected_election_code}&"
                        f"district_id={selected_district_code}&"
                        f"mnc_id={selected_muni_code}&"
                        f"ward_id={ward['code']}&"
                        f"circle_id=0&"
                        f"part_no={p_no}"
                    )
                    
                    filename = f"voterlist_ward{ward['code']}_part{p_no}.pdf"
                    
                    results.append({
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Election": election_options.get(selected_election_code, selected_election_code),
                        "District": district_options.get(selected_district_code, selected_district_code),
                        "Municipality": muni_options.get(selected_muni_code, selected_muni_code),
                        "Ward Name": ward['name'],
                        "Ward Code": ward['code'],
                        "AC Part No": p_no,
                        "Status": "Available",
                        "Link": download_url,
                        "Filename": filename
                    })
            
            progress_bar.progress((i + 1) / total)
            time.sleep(0.05) # Yield to UI
            
        st.session_state['download_results'] = results
        status_text.text("Processing Complete!")
        progress_bar.empty()

# Display Results
if st.session_state['download_results']:
    df = pd.DataFrame(st.session_state['download_results'])
    st.success(f"Found {len(df)} records.")
    
    # Display Data with clickable links
    # Create a copy with clickable links for display
    df_display = df.copy()
    if 'Link' in df_display.columns:
        df_display['Link'] = df_display['Link'].apply(lambda x: f'<a href="{x}" target="_blank">📥 Download</a>')
    
    st.markdown("**Tip:** Click the links below to download PDFs in your browser (most reliable method)")
    st.markdown(df_display.to_html(escape=False, index=False), unsafe_allow_html=True)
    
    # Generate HTML file with all clickable links
    st.markdown("---")
    st.subheader("📄 Download Options")
    
    # Create HTML file with clickable links
    html_content = f'''
<!DOCTYPE html>
<html>
<head>
    <title>Voter List PDF Links - {datetime.now().strftime("%Y-%m-%d %H:%M")}</title>
    <style>
        body {{ font-family: Arial, sans-serif; padding: 20px; }}
        h1 {{ color: #333; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #4CAF50; color: white; }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
        a {{ color: #1a73e8; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .download-btn {{ background: #4CAF50; color: white; padding: 5px 10px; border-radius: 3px; }}
    </style>
</head>
<body>
    <h1>Voter List PDF Download Links</h1>
    <p>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
    <p><strong>Instructions:</strong> Click any link to download the PDF. Make sure you're logged into the TSEC website in the same browser first.</p>
    <table>
        <tr>
            <th>#</th>
            <th>Ward</th>
            <th>Part No</th>
            <th>Download Link</th>
        </tr>
'''
    for idx, row in df.iterrows():
        html_content += f'''        <tr>
            <td>{idx + 1}</td>
            <td>{row.get('Ward Name', row.get('Ward Code', 'N/A'))}</td>
            <td>{row.get('AC Part No', 'N/A')}</td>
            <td><a href="{row['Link']}" target="_blank" class="download-btn">📥 Download PDF</a></td>
        </tr>
'''
    html_content += '''    </table>
</body>
</html>'''
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        st.download_button(
            label="📄 Download HTML with Links",
            data=html_content,
            file_name=f"voter_pdf_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
            mime="text/html",
            help="Download an HTML file with all clickable links. Open in your browser to download PDFs."
        )
    with col_h2:
        st.info("💡 Open the HTML file in your browser and click links to download PDFs")
    
    st.markdown("---")
    st.subheader("📥 Bulk Download Options")
    
    col_d1, col_d2 = st.columns([3, 1])
    with col_d1:
        # Default folder name based on selection
        default_folder = f"voter_pdfs_{selected_muni_code}_ward{str(selected_muni_code)}" 
        target_folder = st.text_input("Save PDFs to folder:", value="extracted_voter_pdfs")
        
    with col_d2:
        st.write("") # Spacer
        st.write("")
        download_clicked = st.button("📥 Download All PDFs", type="primary")
        
    if download_clicked:
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)
            
        progress_bar = st.progress(0)
        status_text = st.empty()
        session = get_session()
        
        success_count = 0
        total_files = len(df)
        
        # CRITICAL: Submit the form first to authorize the session for downloads
        status_text.text("Authorizing session with server...")
        try:
            first_row = df.iloc[0]
            form_url = f"{BASE_URL}/slNoWardWiseVoterlisturbanMapped.do"
            form_data = {
                'mode': 'getWardWiseData',
                'property(election_id)': first_row.get('Election Code', selected_election_code) if 'Election Code' in df.columns else selected_election_code,
                'property(district_id)': first_row.get('District Code', selected_district_code) if 'District Code' in df.columns else selected_district_code,
                'property(municipality_id)': first_row.get('Municipality Code', selected_muni_code) if 'Municipality Code' in df.columns else selected_muni_code,
                'property(ward_id)': first_row['Ward Code'],
                'property(part_no)': first_row['AC Part No']
            }
            session.headers.update({
                'Content-Type': 'application/x-www-form-urlencoded',
                'Referer': form_url
            })
            auth_response = session.post(form_url, data=form_data, verify=False, timeout=60)
            if auth_response.status_code == 200:
                status_text.text("Session authorized. Starting downloads...")
            else:
                st.warning(f"Authorization returned status {auth_response.status_code}. Downloads may fail.")
            # Remove Content-Type for subsequent GET requests
            if 'Content-Type' in session.headers:
                del session.headers['Content-Type']
        except Exception as e:
            st.warning(f"Session authorization step failed: {e}. Proceeding anyway...")

        # Ensure 'Filename' column exists before loop (backward compatibility)
        if 'Filename' not in df.columns:
            df['Filename'] = df.apply(lambda x: f"voterlist_ward{x['Ward Code']}_part{x['AC Part No']}.pdf", axis=1)

        for index, row in df.iterrows():
            filename = row.get('Filename')
            if not filename:
                filename = f"voterlist_ward{row['Ward Code']}_part{row['AC Part No']}.pdf"

            status_text.text(f"Downloading {filename} ({index+1}/{total_files})...")
            
            # Use the link constructed earlier, or reconstruct parameters
            # The 'Link' column has the full URL
            pdf_url = row['Link']
            
            try:
                # Use the session that has the cookies!
                response = session.get(pdf_url, stream=True, timeout=60)
                
                if response.status_code == 200 and 'application/pdf' in response.headers.get('Content-Type', ''):
                    file_path = os.path.join(target_folder, filename)
                    with open(file_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    success_count += 1
                else:
                    st.warning(f"Failed to download {filename}: Status {response.status_code}")
                    # Log error if not PDF
                    if 'application/pdf' not in response.headers.get('Content-Type', ''):
                        st.error(f"Server returned non-PDF content for {filename}")
                        
            except Exception as e:
                st.error(f"Error downloading {filename}: {e}")
                
            progress_bar.progress((index + 1) / total_files)
            
        status_text.success(f"Download Complete! Saved {success_count}/{total_files} files to '{os.path.abspath(target_folder)}'")
        progress_bar.empty()
        
    # Excel Download
    try:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Drop the Link column for cleaner Excel
            df_export = df.drop(columns=['Link'], errors='ignore')
            df_export.to_excel(writer, index=False, sheet_name='VoterData')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📊 Download Excel Summary",
            data=excel_data,
            file_name=f'voter_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='download-excel'
        )
    except Exception as e:
        st.error(f"Excel generation failed: {e}")

# Sidebar Logging
st.sidebar.title("Connection Logs")
if st.button("Clear Logs", key="clear_logs"):
    st.session_state['logs'] = []

if 'logs' in st.session_state:
    for msg in st.session_state['logs']:
        st.sidebar.text(msg)
