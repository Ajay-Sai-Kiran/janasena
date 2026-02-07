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
import glob

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
            election_select = soup.find('select', {'id': 'election_id'})
            if election_select:
                for opt in election_select.find_all('option'):
                    val = opt.get('value')
                    if val and val != '0':
                        elections.append({'id': val, 'name': opt.text.strip()})
            
            # Districts
            districts = []
            district_select = soup.find('select', {'id': 'district_id'})
            if district_select:
                for opt in district_select.find_all('option'):
                    val = opt.get('value')
                    if val and val != '0':
                        districts.append({'id': val, 'name': opt.text.strip()})
                        
            return elections, districts
            
    except Exception as e:
        st.error(f"Error fetching initial data: {e}")
        log_request(url, f"Error: {e}")
        
        st.warning("⚠️ Connection failed. Using fallback data (Nizamabad / 2026 Election).")
        fallback_elections = [{'id': '186', 'name': 'ORDINARY ELECTIONS TO MUNICIPALITIES AND MUNICIPAL CORPORATIONS, 2026'}]
        fallback_districts = [{'id': '05', 'name': 'Nizamabad'}]
        return fallback_elections, fallback_districts

    return [], []

def fetch_municipalities(district_code):
    session = get_session()
    url = f"{BASE_URL}/wardwisevoterlisturban.do"
    params = {'mode': 'getMunicipality', 'district_id': district_code}
    log_request(url, params)
    try:
        response = session.post(url, params=params, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
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
    url = f"{BASE_URL}/wardwisevoterlisturban.do"
    params = {'mode': 'getWard', 'district_id': district_code, 'municipality_id': municipality_code}
    log_request(url, params)
    try:
        response = session.post(url, params=params, timeout=30)
        log_request(url, f"Status: {response.status_code}")
        if response.status_code == 200:
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
    election_options = {e['id']: e['name'] for e in st.session_state['elections']}
    idx = 0
    selected_election_code = st.selectbox(
        "Election", 
        options=list(election_options.keys()), 
        format_func=lambda x: election_options.get(x, x),
        index=idx
    )

    district_options = {d['id']: d['name'] for d in st.session_state['districts']}
    selected_district_code = st.selectbox(
        "District", 
        options=list(district_options.keys()), 
        format_func=lambda x: district_options.get(x, x)
    )

    if 'last_district' not in st.session_state:
        st.session_state['last_district'] = None
        
    if selected_district_code != st.session_state['last_district']:
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
    if 'last_muni' not in st.session_state:
        st.session_state['last_muni'] = None
        
    if selected_muni_code != st.session_state['last_muni']:
         if selected_muni_code:
            with st.spinner("Fetching Wards..."):
                wards = fetch_wards(selected_district_code, selected_muni_code)
                st.session_state['wards'] = wards
                st.session_state['last_muni'] = selected_muni_code
             
    ward_options = {w['id']: w['name'] for w in st.session_state['wards']}
    
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
            time.sleep(0.05)
            
        st.session_state['download_results'] = results
        status_text.text("Processing Complete!")
        progress_bar.empty()

# Display Results
if st.session_state['download_results']:
    df = pd.DataFrame(st.session_state['download_results'])
    st.success(f"Found {len(df)} records.")
    
    df_display = df.copy()
    if 'Link' in df_display.columns:
        df_display['Link'] = df_display['Link'].apply(lambda x: f'<a href="{x}" target="_blank">📥 Download</a>')
    
    st.markdown("**Tip:** Click the links below to download PDFs in your browser (most reliable method)")
    st.markdown(df_display.to_html(escape=False, index=False), unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("📄 Download Options")
    
    html_content = f'''<!DOCTYPE html>
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
        .download-btn {{ background: #4CAF50; color: white; padding: 5px 10px; border-radius: 3px; }}
    </style>
</head>
<body>
    <h1>Voter List PDF Download Links</h1>
    <p>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
    <table>
        <tr><th>#</th><th>Ward</th><th>Part No</th><th>Download Link</th></tr>
'''
    for idx, row in df.iterrows():
        html_content += f'''<tr><td>{idx + 1}</td><td>{row.get('Ward Name', 'N/A')}</td><td>{row.get('AC Part No', 'N/A')}</td><td><a href="{row['Link']}" target="_blank" class="download-btn">📥 Download PDF</a></td></tr>'''
    html_content += '''</table></body></html>'''
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        st.download_button(
            label="📄 Download HTML with Links",
            data=html_content,
            file_name=f"voter_pdf_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
            mime="text/html"
        )
    with col_h2:
        st.info("💡 Open the HTML file in your browser and click links to download PDFs")
    
    # Excel Download
    try:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
    
    # --- Automatic Download & Merge Section ---
    st.markdown("---")
    st.subheader("🚀 Auto-Download & Merge All PDFs")
    st.markdown("Automatically download all PDFs and merge them into a single file using browser automation.")
    
    download_dir = st.text_input(
        "Download folder:", 
        value=os.path.join(os.path.expanduser("~"), "Downloads", "voter_pdfs"),
        key="auto_download_dir"
    )
    
    if st.button("🚀 Download All & Merge into Single PDF", type="primary", key="auto_merge_btn"):
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
        except ImportError:
            st.error("❌ Selenium not installed. Run: `pip install selenium webdriver-manager`")
            st.stop()
        
        try:
            from PyPDF2 import PdfMerger
        except ImportError:
            st.error("❌ PyPDF2 not installed. Run: `pip install PyPDF2`")
            st.stop()
        
        # Create download directory
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        
        st.info("🌐 Starting browser for automated downloads...")
        
        # Setup Chrome with PDF download preferences
        options = Options()
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--ignore-certificate-errors")
        
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
        except:
            try:
                driver = webdriver.Chrome(options=options)
            except Exception as e:
                st.error(f"Failed to start Chrome: {e}")
                st.info("Make sure Chrome and ChromeDriver are installed.")
                st.stop()
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            total_parts = len(df)
            for idx, row in df.iterrows():
                part_no = row.get('AC Part No', idx)
                ward_name = row.get('Ward Name', row.get('Ward Code', 'Unknown'))
                status_text.text(f"Downloading Ward {ward_name} Part {part_no} ({idx+1}/{total_parts})...")
                
                pdf_url = row['Link']
                driver.get(pdf_url)
                time.sleep(3)  # Wait for download
                
                progress_bar.progress((idx + 1) / total_parts)
            
            status_text.text("Downloads complete. Closing browser...")
            driver.quit()
            
            # Wait for any pending downloads
            time.sleep(2)
            
            # Find all PDFs in download folder
            pdf_files = sorted(glob.glob(os.path.join(download_dir, "*.pdf")))
            
            if pdf_files:
                st.success(f"✅ Found {len(pdf_files)} PDF files in download folder")
                
                # Merge PDFs
                status_text.text("Merging PDFs into single file...")
                merger = PdfMerger()
                merged_count = 0
                
                for pdf_file in pdf_files:
                    try:
                        merger.append(pdf_file)
                        merged_count += 1
                    except Exception as e:
                        st.warning(f"⚠️ Failed to add {os.path.basename(pdf_file)}: {e}")
                
                # Get ward name for output filename
                first_ward = df.iloc[0].get('Ward Name', df.iloc[0].get('Ward Code', 'unknown'))
                output_file = os.path.join(download_dir, f"merged_ward_{first_ward}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                merger.write(output_file)
                merger.close()
                
                st.success(f"✅ Merged {merged_count} PDFs into: {output_file}")
                
                # Offer download button
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Merged PDF",
                        data=f.read(),
                        file_name=os.path.basename(output_file),
                        mime="application/pdf",
                        key="download-auto-merged"
                    )
            else:
                st.warning("⚠️ No PDF files found in download folder.")
                st.info("💡 The website may have returned HTML pages instead of PDFs. Try the manual method: download HTML file and click links in your browser.")
                
        except Exception as e:
            st.error(f"Error during automation: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            try:
                driver.quit()
            except:
                pass
        
        progress_bar.empty()
        status_text.empty()

# --- PDF to Excel Conversion & Merge Section ---
st.markdown("---")
st.header("📄 PDF Tools")

tab_merge, tab_extract = st.tabs(["📎 Merge PDFs", "📊 PDF to Excel"])

with tab_merge:
    st.markdown("Upload multiple voter list PDFs to merge them into a single PDF file.")
    
    merge_files = st.file_uploader(
        "Upload PDFs to merge",
        type=['pdf'],
        accept_multiple_files=True,
        key="merge_uploader"
    )
    
    if merge_files:
        st.info(f"📁 {len(merge_files)} file(s) selected")
        
        if st.button("📎 Merge into Single PDF", type="primary", key="merge_btn"):
            try:
                from PyPDF2 import PdfMerger
                from io import BytesIO
                
                merger = PdfMerger()
                for pdf in merge_files:
                    pdf.seek(0)
                    merger.append(pdf)
                
                output = BytesIO()
                merger.write(output)
                merger.close()
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Merged PDF",
                    data=output.getvalue(),
                    file_name=f"merged_voterlist_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    key="download-merged"
                )
                st.success(f"✅ Merged {len(merge_files)} PDFs successfully!")
            except ImportError:
                st.error("❌ PyPDF2 not installed. Run: `pip install PyPDF2`")
            except Exception as e:
                st.error(f"Merge failed: {e}")

with tab_extract:
    st.markdown("Upload voter list PDFs to extract voter data into Excel format.")
    
    extract_files = st.file_uploader(
        "Upload PDFs to extract data",
        type=['pdf'],
        accept_multiple_files=True,
        key="extract_uploader"
    )
    
    if extract_files:
        st.info(f"📁 {len(extract_files)} file(s) selected")
        
        if st.button("🔄 Convert PDFs to Excel", type="primary", key="extract_btn"):
            try:
                import pdfplumber
            except ImportError:
                st.error("❌ pdfplumber not installed. Run: `pip install pdfplumber`")
                st.stop()
            
            all_voters = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for file_idx, uploaded_file in enumerate(extract_files):
                status_text.text(f"Processing {uploaded_file.name} ({file_idx + 1}/{len(extract_files)})...")
                
                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page_num, page in enumerate(pdf.pages):
                            text = page.extract_text()
                            if not text:
                                continue
                            
                            # Extract ward from filename
                            ward_match = re.search(r'ward[-_]?(\d+)', uploaded_file.name, re.IGNORECASE)
                            ward_from_filename = ward_match.group(1) if ward_match else 'Unknown'
                            
                            lines = text.split('\n')
                            current_voter = {}
                            
                            for line in lines:
                                line = line.strip()
                                
                                # Match AC No.-PS No.-SLNo pattern
                                ac_match = re.search(r'A\.?C\.?\s*No\.?.*?PS\s*No\.?.*?SL\.?\s*No\.?.*?:\s*(\d+)\s*[-–]\s*(\d+)\s*[-–]\s*(\d+)', line, re.IGNORECASE)
                                if ac_match:
                                    if current_voter and current_voter.get('Name'):
                                        all_voters.append(current_voter)
                                    current_voter = {
                                        'Source File': uploaded_file.name,
                                        'Ward': ward_from_filename,
                                        'AC No': ac_match.group(1),
                                        'PS No': ac_match.group(2),
                                        'SL No': ac_match.group(3)
                                    }
                                    continue
                                
                                # Match Name
                                name_match = re.match(r'^Name\s*[:.]?\s*(.+)$', line, re.IGNORECASE)
                                if name_match and current_voter:
                                    current_voter['Name'] = name_match.group(1).strip()
                                    continue
                                
                                # Match Father/Husband Name
                                father_match = re.match(r'^(Father|Husband)\s*(Name)?\s*[:.]?\s*(.+)$', line, re.IGNORECASE)
                                if father_match and current_voter:
                                    current_voter['Father/Husband Name'] = father_match.group(3).strip()
                                    continue
                                
                                # Match Age and Sex
                                age_match = re.search(r'Age\s*[:.]?\s*(\d+)', line, re.IGNORECASE)
                                sex_match = re.search(r'Sex\s*[:.]?\s*([MF])', line, re.IGNORECASE)
                                if age_match and current_voter:
                                    current_voter['Age'] = age_match.group(1)
                                if sex_match and current_voter:
                                    current_voter['Sex'] = sex_match.group(1)
                                
                                # Match Door No
                                door_match = re.match(r'^Door\s*No\.?\s*[:.]?\s*(.+)$', line, re.IGNORECASE)
                                if door_match and current_voter:
                                    current_voter['Door No'] = door_match.group(1).strip()
                                    continue
                                
                                # Match EPIC No
                                epic_match = re.match(r'^EPIC\s*No\.?\s*[:.]?\s*([A-Z0-9]+)', line, re.IGNORECASE)
                                if epic_match and current_voter:
                                    current_voter['EPIC No'] = epic_match.group(1).strip()
                                    continue
                            
                            if current_voter and current_voter.get('Name'):
                                all_voters.append(current_voter)
                                
                except Exception as e:
                    st.warning(f"⚠️ Error processing {uploaded_file.name}: {e}")
                
                progress_bar.progress((file_idx + 1) / len(extract_files))
            
            progress_bar.empty()
            status_text.empty()
            
            if all_voters:
                df_voters = pd.DataFrame(all_voters)
                
                cols = ['Source File', 'Ward', 'AC No', 'PS No', 'SL No', 'Name', 
                       'Father/Husband Name', 'Age', 'Sex', 'Door No', 'EPIC No']
                cols = [c for c in cols if c in df_voters.columns]
                df_voters = df_voters[cols]
                
                st.success(f"✅ Extracted {len(df_voters)} voter records from {len(extract_files)} PDF(s)")
                
                # Group by ward for separate sheets
                wards = df_voters['Ward'].unique()
                st.info(f"📋 Found {len(wards)} ward(s): {', '.join(sorted([str(w) for w in wards]))}")
                
                st.dataframe(df_voters.head(100), use_container_width=True)
                if len(df_voters) > 100:
                    st.caption(f"Showing first 100 of {len(df_voters)} records")
                
                # Excel with separate sheets per ward
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_voters.to_excel(writer, index=False, sheet_name='All Voters')
                    for ward in sorted(wards):
                        ward_df = df_voters[df_voters['Ward'] == ward]
                        sheet_name = f"Ward_{ward}"[:31]
                        ward_df.to_excel(writer, index=False, sheet_name=sheet_name)
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=output.getvalue(),
                    file_name=f"voter_data_extracted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download-extracted"
                )
            else:
                st.warning("⚠️ No voter data could be extracted.")

# Sidebar Logging
st.sidebar.title("Connection Logs")
if st.button("Clear Logs", key="clear_logs"):
    st.session_state['logs'] = []

if 'logs' in st.session_state:
    for msg in st.session_state['logs']:
        st.sidebar.text(msg)
