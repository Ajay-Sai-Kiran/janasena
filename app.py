import streamlit as st
import pandas as pd
import requests
import json
import time
from datetime import datetime
import os
from bs4 import BeautifulSoup
import re
import glob
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def create_session():
    """Returns a configured requests.Session object with retries."""
    session = requests.Session()
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Referer': 'https://urban2025.tsec.gov.in/slNoWardWiseVoterlisturbanMapped.do',
        'Origin': 'https://urban2025.tsec.gov.in',
    }
    
    session.headers.update(headers)
    session.verify = False
    
    retry_strategy = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "POST", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    
    return session

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
        session = create_session()  # Use embedded function
        try:
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
    
    st.markdown("**📥 Click buttons below to download PDFs directly:**")
    
    # Create individual download buttons for each PDF
    session = get_session()
    
    # Show in expandable sections by ward
    wards = df['Ward Name'].unique() if 'Ward Name' in df.columns else df['Ward Code'].unique()
    
    for ward in wards:
        ward_df = df[df['Ward Name'] == ward] if 'Ward Name' in df.columns else df[df['Ward Code'] == ward]
        
        with st.expander(f"📁 {ward} ({len(ward_df)} parts)", expanded=True):
            cols = st.columns(5)  # 5 buttons per row
            
            for idx, (_, row) in enumerate(ward_df.iterrows()):
                col_idx = idx % 5
                part_no = row.get('AC Part No', 'X')
                filename = f"ward{row.get('Ward Code', 'X')}_part{part_no}.pdf"
                
                with cols[col_idx]:
                    if st.button(f"📥 Part {part_no}", key=f"dl_{ward}_{part_no}"):
                        try:
                            with st.spinner(f"Downloading..."):
                                response = session.get(row['Link'], timeout=60)
                                if response.status_code == 200 and 'pdf' in response.headers.get('Content-Type', '').lower():
                                    st.download_button(
                                        label=f"💾 Save {filename}",
                                        data=response.content,
                                        file_name=filename,
                                        mime="application/pdf",
                                        key=f"save_{ward}_{part_no}"
                                    )
                                else:
                                    st.error(f"Failed: Not PDF")
                        except Exception as e:
                            st.error(f"Error: {str(e)[:30]}")
    
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
        filename = f"voterlist_ward{row.get('Ward Code', 'X')}_part{row.get('AC Part No', 'X')}.pdf"
        html_content += f'''<tr><td>{idx + 1}</td><td>{row.get('Ward Name', 'N/A')}</td><td>{row.get('AC Part No', 'N/A')}</td><td><a href="{row['Link']}" download="{filename}" class="download-btn">📥 Download PDF</a></td></tr>'''
    html_content += '''</table>
    <script>
    // Auto-click all download links with delay
    function downloadAll() {
        const links = document.querySelectorAll('a.download-btn');
        let delay = 0;
        links.forEach((link, i) => {
            setTimeout(() => {
                link.click();
                document.getElementById('status').innerText = 'Downloading ' + (i+1) + '/' + links.length;
            }, delay);
            delay += 2000; // 2 second delay between downloads
        });
    }
    </script>
    <br><button onclick="downloadAll()" style="background:#4CAF50;color:white;padding:10px 20px;font-size:16px;cursor:pointer;">📥 Download All PDFs (Auto-Click)</button>
    <span id="status"></span>
</body></html>'''
    
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
    
    # --- In-Memory Download & Merge Section (Works on deployed apps) ---
    st.markdown("---")
    st.subheader("🚀 Auto-Download & Merge All PDFs")
    st.markdown("Download all PDFs and merge them into a single file. Works on deployed apps!")
    
    st.warning("⚠️ **Note:** The TSEC website may block automated downloads. If this fails, use the 'Upload & Merge' option below instead.")
    
    if st.button("🚀 Download All & Merge into Single PDF", type="primary", key="auto_merge_btn"):
        try:
            from PyPDF2 import PdfMerger
        except ImportError:
            st.error("❌ PyPDF2 not installed. Run: `pip install PyPDF2`")
            st.stop()
        
        from io import BytesIO
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        session = get_session()
        pdf_buffers = []
        failed_downloads = []
        
        # CRITICAL: Authorize the session first by submitting the form
        status_text.text("Authorizing session with TSEC server...")
        try:
            first_row = df.iloc[0]
            auth_url = f"{BASE_URL}/slNoWardWiseVoterlisturbanMapped.do"
            
            # First, visit the main page to get fresh cookies
            session.get(auth_url, timeout=30)
            
            # Submit form to authorize PDF downloads
            form_data = {
                'mode': 'getWardWiseData',
                'property(election_id)': first_row.get('Election', '186'),
                'property(district_id)': first_row.get('District', '05'),
                'property(municipality_id)': first_row.get('Municipality', '1'),
                'property(ward_id)': first_row.get('Ward Code', '1'),
                'property(part_no)': first_row.get('AC Part No', '1')
            }
            
            session.headers.update({
                'Content-Type': 'application/x-www-form-urlencoded',
                'Referer': auth_url
            })
            
            auth_response = session.post(auth_url, data=form_data, timeout=60)
            
            if auth_response.status_code == 200:
                status_text.text("✅ Session authorized. Starting downloads...")
            else:
                st.warning(f"Authorization returned {auth_response.status_code}. Downloads may fail.")
            
            # Remove Content-Type for subsequent GET requests
            if 'Content-Type' in session.headers:
                del session.headers['Content-Type']
                
        except Exception as e:
            st.warning(f"Session auth failed: {e}. Trying downloads anyway...")
        
        total_parts = len(df)
        for idx, row in df.iterrows():
            part_no = row.get('AC Part No', idx)
            ward_name = row.get('Ward Name', row.get('Ward Code', 'Unknown'))
            status_text.text(f"Downloading Ward {ward_name} Part {part_no} ({idx+1}/{total_parts})...")
            
            pdf_url = row['Link']
            
            try:
                response = session.get(pdf_url, timeout=60, stream=True)
                
                if response.status_code == 200:
                    content_type = response.headers.get('Content-Type', '')
                    
                    if 'pdf' in content_type.lower():
                        pdf_buffer = BytesIO(response.content)
                        pdf_buffers.append((f"ward{ward_name}_part{part_no}.pdf", pdf_buffer))
                    else:
                        failed_downloads.append(f"Part {part_no}: Not a PDF (got {content_type})")
                else:
                    failed_downloads.append(f"Part {part_no}: HTTP {response.status_code}")
            except Exception as e:
                failed_downloads.append(f"Part {part_no}: {str(e)[:50]}")
            
            progress_bar.progress((idx + 1) / total_parts)
        
        progress_bar.empty()
        status_text.empty()
        
        if pdf_buffers:
            st.success(f"✅ Successfully downloaded {len(pdf_buffers)}/{total_parts} PDFs")
            
            if failed_downloads:
                with st.expander(f"⚠️ {len(failed_downloads)} downloads failed"):
                    for fail in failed_downloads:
                        st.text(fail)
            
            # Merge PDFs
            try:
                merger = PdfMerger()
                for filename, pdf_buffer in pdf_buffers:
                    pdf_buffer.seek(0)
                    merger.append(pdf_buffer)
                
                merged_output = BytesIO()
                merger.write(merged_output)
                merger.close()
                merged_output.seek(0)
                
                first_ward = df.iloc[0].get('Ward Name', df.iloc[0].get('Ward Code', 'unknown'))
                
                st.download_button(
                    label=f"📥 Download Merged PDF ({len(pdf_buffers)} parts)",
                    data=merged_output.getvalue(),
                    file_name=f"merged_ward_{first_ward}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    key="download-auto-merged"
                )
                
                st.success(f"✅ Merged {len(pdf_buffers)} PDFs successfully!")
                
            except Exception as e:
                st.error(f"Merge failed: {e}")
        else:
            st.error("❌ No PDFs could be downloaded.")
            st.info("💡 The website may require browser authentication. Try the manual method:")
            st.markdown("1. Download the **HTML with Links** file above")
            st.markdown("2. Open it in your browser and click each link to download PDFs")
            st.markdown("3. Use the **Upload & Merge** section below to merge them")

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
