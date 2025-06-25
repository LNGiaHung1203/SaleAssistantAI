# -*- coding: utf-8 -*-
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import requests
import time
import random
import re
import os
import json
from datetime import datetime
from find_linkedin import find_linkedin_url

# === USER CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'ggsheetapikey.json'
SPREADSHEET_ID = '1EnqcDxCPnm7CSfAntipxPpC9i6tEYk5nXVMMud4lOek'
COMPANY_NAME_COL = 'company_name'
LINKEDIN_URL_COL = 'LinkedIn URL'
LINKEDIN_POSTS_COL = 'linkedin_posts'
LINKEDIN_POSTS_FILE_COL = 'linkedin_posts_file'
LINKEDIN_POSTS_JSON_COL = 'linkedin_posts_json'
BRAVE_PATH = r'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'
DATA_DIR = 'company_data'

# === ANTI-DETECTION CONFIGURATION ===
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/120.0.0.0'
]

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
for col in [LINKEDIN_POSTS_COL, LINKEDIN_POSTS_FILE_COL, LINKEDIN_POSTS_JSON_COL]:
    if col not in header:
        worksheet.update_cell(1, len(header) + 1, col)
        header.append(col)

linkedin_posts_col_idx = header.index(LINKEDIN_POSTS_COL) + 1
linkedin_posts_file_col_idx = header.index(LINKEDIN_POSTS_FILE_COL) + 1
linkedin_posts_json_col_idx = header.index(LINKEDIN_POSTS_JSON_COL) + 1
linkedin_col_idx = header.index(LINKEDIN_URL_COL) + 1
company_col_idx = header.index(COMPANY_NAME_COL) + 1

# === SETUP DATA DIR ===
os.makedirs(DATA_DIR, exist_ok=True)

def setup_driver():
    """Setup Chrome driver with anti-detection measures for no-login scraping"""
    chrome_options = Options()
    chrome_options.binary_location = BRAVE_PATH
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-plugins')
    chrome_options.add_argument('--disable-images')  # Faster loading
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-web-security')
    chrome_options.add_argument('--allow-running-insecure-content')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Random user agent
    user_agent = random.choice(USER_AGENTS)
    chrome_options.add_argument(f'--user-agent={user_agent}')
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def random_delay(min_seconds=2, max_seconds=6):
    """Add random delay to prevent detection"""
    delay = random.uniform(min_seconds, max_seconds)
    print(f"Waiting {delay:.2f} seconds...")
    time.sleep(delay)

def scroll_page(driver, scroll_pause_time=2, max_scrolls=5):
    """Scroll page to load more content"""
    last_height = driver.execute_script("return document.body.scrollHeight")
    scroll_count = 0
    
    while scroll_count < max_scrolls:
        # Scroll down
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        random_delay(scroll_pause_time, scroll_pause_time + 2)
        
        # Calculate new scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        scroll_count += 1

def extract_linkedin_posts_no_login(linkedin_url, company_name):
    """Extract LinkedIn posts without login using specific HTML class"""
    posts_data = []
    posts_text = ''
    posts_file = None
    posts_json_file = None
    
    if not linkedin_url or 'linkedin.com/company/' not in linkedin_url:
        return posts_data, posts_text, posts_file, posts_json_file
    
    # Normalize URL
    base_url = linkedin_url.split('?')[0].rstrip('/')
    posts_url = base_url + '/posts/'
    
    driver = None
    try:
        driver = setup_driver()
        
        # Navigate to posts page
        print(f"Navigating to posts page: {posts_url}")
        driver.get(posts_url)
        random_delay(3, 6)
        
        # Wait for page to load
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        except TimeoutException:
            print("Page load timeout")
            return posts_data, posts_text, posts_file, posts_json_file
        
        # Scroll to load more posts
        scroll_page(driver, scroll_pause_time=2, max_scrolls=3)
        
        # Extract posts using the specific class you mentioned
        post_selectors = [
            # Primary selector - the class you specified
            'div.core-section-container__content.break-words',
            
            # Alternative selectors for different post types
            'div.feed-shared-update-v2__description',
            'div.update-components-text',
            'div.feed-shared-text',
            'span.break-words',
            'div.feed-shared-text__text-view',
            
            # Generic post content selectors
            'div[class*="core-section-container"]',
            'div[class*="break-words"]',
            'div[class*="feed-shared"]',
            'div[class*="update-components"]'
        ]
        
        posts_found = []
        for selector in post_selectors:
            try:
                posts = driver.find_elements(By.CSS_SELECTOR, selector)
                if posts:
                    print(f"Found {len(posts)} posts with selector: {selector}")
                    posts_found.extend(posts)
                    # If we found posts with the primary selector, use only those
                    if 'core-section-container__content break-words' in selector:
                        posts_found = posts
                        break
            except Exception as e:
                print(f"Error with selector {selector}: {e}")
                continue
        
        # Remove duplicates while preserving order
        seen_texts = set()
        unique_posts = []
        for post in posts_found:
            try:
                post_text = post.text.strip()
                if post_text and len(post_text) > 10 and post_text not in seen_texts:
                    seen_texts.add(post_text)
                    unique_posts.append(post)
            except Exception:
                continue
        
        # Extract post content
        for i, post in enumerate(unique_posts[:30]):  # Limit to 30 posts
            try:
                post_text = post.text.strip()
                if post_text and len(post_text) > 10:  # Filter out very short posts
                    post_data = {
                        'index': i + 1,
                        'text': post_text,
                        'length': len(post_text),
                        'timestamp': datetime.now().isoformat(),
                        'selector_used': 'core-section-container__content break-words'
                    }
                    posts_data.append(post_data)
            except Exception as e:
                print(f"Error extracting post {i}: {e}")
        
        # Create text file
        if posts_data:
            posts_text = '\n\n--- POST ---\n\n'.join([f"Post {p['index']}:\n{p['text']}" for p in posts_data])
            posts_file = os.path.join(DATA_DIR, f"{company_name.replace(' ', '_')}_linkedin_posts.txt")
            with open(posts_file, 'w', encoding='utf-8') as f:
                f.write(posts_text)
            
            # Create JSON file
            posts_json_file = os.path.join(DATA_DIR, f"{company_name.replace(' ', '_')}_linkedin_posts.json")
            with open(posts_json_file, 'w', encoding='utf-8') as f:
                json.dump(posts_data, f, indent=2, ensure_ascii=False)
        
        print(f"Extracted {len(posts_data)} posts for {company_name}")
        
    except Exception as e:
        print(f'Error scraping LinkedIn posts for {linkedin_url}: {e}')
    finally:
        if driver:
            driver.quit()
    
    return posts_data, posts_text, posts_file, posts_json_file

def main():
    """Main function to process all companies without login"""
    print("Starting LinkedIn posts scraper (NO LOGIN REQUIRED)...")
    print("Using HTML class: core-section-container__content break-words")
    print("=" * 60)
    
    # Process each company
    for i, row in enumerate(records, start=2):
        company_name = row.get(COMPANY_NAME_COL)
        linkedin_url = row.get(LINKEDIN_URL_COL)
        
        # Skip if no LinkedIn URL
        if not linkedin_url or linkedin_url == 'NOT FOUND':
            print(f"Skipping {company_name}: No LinkedIn URL")
            continue
        
        # Check if already processed
        existing_posts = row.get(LINKEDIN_POSTS_COL)
        if existing_posts:
            print(f"Skipping {company_name}: Already has posts data")
            continue
        
        print(f"\nProcessing {company_name} ({linkedin_url})")
        
        # Extract posts without login
        posts_data, posts_text, posts_file, posts_json_file = extract_linkedin_posts_no_login(linkedin_url, company_name)
        
        # Update spreadsheet
        if posts_text:
            worksheet.update_cell(i, linkedin_posts_col_idx, posts_text[:500])  # Preview
        if posts_file:
            worksheet.update_cell(i, linkedin_posts_file_col_idx, posts_file)
        if posts_json_file:
            worksheet.update_cell(i, linkedin_posts_json_col_idx, posts_json_file)
        
        # Random delay between companies
        random_delay(5, 12)
    
    print("\nLinkedIn posts scraping completed (No login required)!")

if __name__ == "__main__":
    main() 