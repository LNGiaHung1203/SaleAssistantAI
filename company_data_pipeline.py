import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import requests
import time
import random
import re
import os
from find_linkedin import find_linkedin_url

# === USER CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'ggsheetapikey.json'
SPREADSHEET_ID = '1EnqcDxCPnm7CSfAntipxPpC9i6tEYk5nXVMMud4lOek'
COMPANY_NAME_COL = 'company_name'
DOMAIN_NAME_COL = 'domain_name'
LINKEDIN_URL_COL = 'LinkedIn URL'
IS_HIRING_COL = 'is_hiring'
NUM_JOBS_COL = 'num_jobs'
LINKEDIN_POSTS_COL = 'linkedin_posts'
PROFILE_HTML_COL = 'profile_html_file'
LINKEDIN_POSTS_FILE_COL = 'linkedin_posts_file'
BRAVE_PATH = r'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'  # Update if needed
LINKEDIN_USERNAME = 'edwardwateryhung@gmail.com'  # <-- Put your LinkedIn email here
LINKEDIN_PASSWORD = 'giahung1232003'           # <-- Put your LinkedIn password here
DATA_DIR = 'company_data'

# === AUTHENTICATE ===
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# === OPEN SHEET ===
sh = gc.open_by_key(SPREADSHEET_ID)
worksheet = sh.sheet1

# === READ DATA ===
records = worksheet.get_all_records()
header = worksheet.row_values(1)
print('Sheet headers:', header)

# Ensure all needed columns exist
for col in [LINKEDIN_URL_COL, IS_HIRING_COL, NUM_JOBS_COL, LINKEDIN_POSTS_COL, PROFILE_HTML_COL, LINKEDIN_POSTS_FILE_COL]:
    if col not in header:
        worksheet.update_cell(1, len(header) + 1, col)
        header.append(col)
linkedin_col_idx = header.index(LINKEDIN_URL_COL) + 1
is_hiring_col_idx = header.index(IS_HIRING_COL) + 1
num_jobs_col_idx = header.index(NUM_JOBS_COL) + 1
linkedin_posts_col_idx = header.index(LINKEDIN_POSTS_COL) + 1
profile_html_col_idx = header.index(PROFILE_HTML_COL) + 1
linkedin_posts_file_col_idx = header.index(LINKEDIN_POSTS_FILE_COL) + 1
domain_col_idx = header.index(DOMAIN_NAME_COL) + 1
company_col_idx = header.index(COMPANY_NAME_COL) + 1

# === SETUP DATA DIR ===
os.makedirs(DATA_DIR, exist_ok=True)

# === SETUP SELENIUM FOR BRAVE ===
chrome_options = Options()
chrome_options.binary_location = BRAVE_PATH
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
driver = webdriver.Chrome(options=chrome_options)

def linkedin_auto_login(driver, username, password):
    driver.get('https://www.linkedin.com/login')
    time.sleep(2)
    user_input = driver.find_element(By.ID, 'username')
    pass_input = driver.find_element(By.ID, 'password')
    user_input.send_keys(username)
    pass_input.send_keys(password)
    pass_input.submit()
    time.sleep(3)

linkedin_auto_login(driver, LINKEDIN_USERNAME, LINKEDIN_PASSWORD)


def download_company_html(domain, company_name):
    if not domain:
        return None
    url = domain if domain.startswith('http') else f'http://{domain}'
    try:
        resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code == 200:
            filename = os.path.join(DATA_DIR, f"{company_name.replace(' ', '_')}_profile.html")
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(resp.text)
            return filename
    except Exception as e:
        print(f"Error downloading {url}: {e}")
    return None


def check_linkedin_jobs_and_posts(linkedin_url, company_name):
    is_hiring = 'No jobs found'
    num_jobs = 0
    posts_text = ''
    posts_file = None
    if not linkedin_url or 'linkedin.com/company/' not in linkedin_url:
        return is_hiring, num_jobs, posts_text, posts_file
    # Normalize URL
    base_url = linkedin_url.split('?')[0].rstrip('/')
    jobs_url = base_url + '/jobs/'
    posts_url = base_url + '/posts/'
    try:
        # Jobs
        driver.get(jobs_url)
        time.sleep(random.uniform(4, 7))
        try:
            headline = driver.find_element(By.CSS_SELECTOR, 'h4.org-jobs-job-search-form-module__headline')
            if headline:
                text = headline.text.strip()
                match = re.search(r'has (\d+) job opening', text)
                if match:
                    is_hiring = 'Yes'
                    num_jobs = int(match.group(1))
                else:
                    is_hiring = text
        except Exception:
            pass
        # Posts
        driver.get(posts_url)
        time.sleep(random.uniform(4, 7))
        try:
            posts = driver.find_elements(By.CSS_SELECTOR, 'div.feed-shared-update-v2__description, div.update-components-text')
            posts_text = '\n\n'.join([p.text for p in posts if p.text.strip()])
            if posts_text:
                posts_file = os.path.join(DATA_DIR, f"{company_name.replace(' ', '_')}_linkedin_posts.txt")
                with open(posts_file, 'w', encoding='utf-8') as f:
                    f.write(posts_text)
        except Exception:
            pass
    except Exception as e:
        print(f'Error scraping LinkedIn for {linkedin_url}: {e}')
    return is_hiring, num_jobs, posts_text, posts_file

# === PROCESS EACH COMPANY ===
for i, row in enumerate(records, start=2):
    company_name = row.get(COMPANY_NAME_COL)
    domain = row.get(DOMAIN_NAME_COL)
    linkedin_url = row.get(LINKEDIN_URL_COL)
    # Step 1: Download and save company homepage HTML
    html_file = row.get(PROFILE_HTML_COL)
    if not html_file and domain:
        print(f"Downloading HTML for: {company_name} ({domain})")
        html_file = download_company_html(domain, company_name)
        if html_file:
            worksheet.update_cell(i, profile_html_col_idx, html_file)
        time.sleep(random.uniform(2, 4))
    # Step 2: Find LinkedIn URL if missing
    if not linkedin_url and company_name:
        print(f"Searching LinkedIn for: {company_name}")
        linkedin_url = find_linkedin_url(company_name)
        worksheet.update_cell(i, linkedin_col_idx, linkedin_url if linkedin_url else 'NOT FOUND')
        print(f"LinkedIn URL: {linkedin_url}")
        time.sleep(random.uniform(4, 8))
    # Step 3: Scrape LinkedIn jobs and posts
    if linkedin_url and linkedin_url != 'NOT FOUND':
        print(f"Scraping LinkedIn for: {linkedin_url}")
        is_hiring, num_jobs, posts_text, posts_file = check_linkedin_jobs_and_posts(linkedin_url, company_name)
        worksheet.update_cell(i, is_hiring_col_idx, is_hiring)
        worksheet.update_cell(i, num_jobs_col_idx, num_jobs)
        if posts_text:
            worksheet.update_cell(i, linkedin_posts_col_idx, posts_text[:500])  # Save a preview
        if posts_file:
            worksheet.update_cell(i, linkedin_posts_file_col_idx, posts_file)
        print(f"is_hiring: {is_hiring}, num_jobs: {num_jobs}, posts_file: {posts_file}")
        time.sleep(random.uniform(4, 8))

# === TODO: LLM PROCESSING ===
# For each company, load the HTML and LinkedIn posts files, send to LLM for extraction and prediction.
# Save LLM results (profile, interests, purchase probability, reasons) to new columns in the sheet.


driver.quit()
print("Done!") 