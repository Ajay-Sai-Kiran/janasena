import requests
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_session():
    """
    Returns a configured requests.Session object with retries.
    """
    session = requests.Session()
    
    # Headers from user browser (updated 2026-02-07)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'en-US,en;q=0.9,en-IN;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        # 'Cookie': 'JSESSIONID=...', # Removed hardcoded cookie to allow fresh session
        'Referer': 'https://urban2025.tsec.gov.in/slNoWardWiseVoterlisturbanMapped.do',
        'Origin': 'https://urban2025.tsec.gov.in',
        'Sec-Ch-Ua': '"Not(A:Brand";v="8", "Chromium";v="144", "Microsoft Edge";v="144"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1'
    }
    
    session.headers.update(headers)
    session.verify = False
    
    # Add Retry Logic
    retry_strategy = Retry(
        total=5, # Increased retries
        connect=5, # Retry connection errors specifically
        read=5,
        backoff_factor=2, # Slower backoff
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "POST", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    
    return session
