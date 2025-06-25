# -*- coding: utf-8 -*-
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import random
import os
import json
from datetime import datetime

# === USER CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'ggsheetapikey.json'
SPREADSHEET_ID = '1EnqcDxCPnm7CSfAntipxPpC9i6tEYk5nXVMMud4lOek'
COMPANY_NAME_COL = 'company_name'
LINKEDIN_URL_COL = 'LinkedIn URL'
LINKEDIN_POSTS_COL = 'linkedin_posts'
LINKEDIN_POSTS_FILE_COL = 'linkedin_posts_file'
BRAVE_PATH = r'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'
DATA_DIR = 'company_data'

# === AUTHENTICATE ===
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# === OPEN SHEET ===
sh = gc.open_by_key(SPREADSHEET_ID)
worksheet = sh.sheet1
records = worksheet.get_all_records()
header = worksheet.row_values(1)

# Ensure columns exist
for col in [LINKEDIN_POSTS_COL, LINKEDIN_POSTS_FILE_COL]:
    if col not in header:
        worksheet.update_cell(1, len(header) + 1, col)
        header.append(col)

# Get column indices
linkedin_posts_col_idx = header.index(LINKEDIN_POSTS_COL) + 1
linkedin_posts_file_col_idx = header.index(LINKEDIN_POSTS_FILE_COL) + 1
linkedin_col_idx = header.index(LINKEDIN_URL_COL) + 1
company_col_idx = header.index(COMPANY_NAME_COL) + 1

# Setup data directory
os.makedirs(DATA_DIR, exist_ok=True)

def setup_driver():
    """Simple driver setup for no-login scraping"""
    chrome_options = Options()
    chrome_options.binary_location = BRAVE_PATH
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-images')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def random_delay(min_seconds=2, max_seconds=5):
    """Random delay to prevent detection"""
    delay = random.uniform(min_seconds, max_seconds)
    print(f"Waiting {delay:.2f} seconds...")
    time.sleep(delay)

def get_linkedin_posts(linkedin_url, company_name):
    """Get LinkedIn posts from main company page (not posts page)"""
    if not linkedin_url or 'linkedin.com/company/' not in linkedin_url:
        return [], '', None
    
    # Use the main company page instead of posts page
    # LinkedIn redirects to login for /posts/ when not authenticated
    company_url = linkedin_url.split('?')[0].rstrip('/')
    
    driver = None
    try:
        driver = setup_driver()
        
        print(f"Getting posts from main company page: {company_url}")
        driver.get(company_url)
        random_delay(3, 6)
        
        # Check if we're redirected to login page
        current_url = driver.current_url
        if 'login' in current_url or 'signup' in current_url:
            print("Redirected to login page. Trying alternative approach...")
            return [], '', None
        
        # Scroll to load more content
        for _ in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            random_delay(2, 4)
        
        # Try to find posts on the main company page
        # Look for posts in the feed section
        post_selectors = [
            # Primary selector - the class you specified
            'div.core-section-container__content.break-words',
            
            # Alternative selectors for posts on company page
            'div.feed-shared-update-v2__description',
            'div.update-components-text',
            'div.feed-shared-text',
            'span.break-words',
            'div.feed-shared-text__text-view',
            
            # Company page specific selectors
            'div[data-test-id="company-post"]',
            'div[class*="feed-shared"]',
            'div[class*="update-components"]',
            'div[class*="core-section"]',
            
            # Generic content selectors
            'div[class*="content"]',
            'div[class*="text"]',
            'span[class*="text"]'
        ]
        
        posts_found = []
        for selector in post_selectors:
            try:
                posts = driver.find_elements(By.CSS_SELECTOR, selector)
                if posts:
                    print(f"Found {len(posts)} elements with selector: {selector}")
                    posts_found.extend(posts)
                    # If we found elements with the primary selector, use only those
                    if 'core-section-container__content break-words' in selector:
                        posts_found = posts
                        break
            except Exception as e:
                print(f"Error with selector {selector}: {e}")
                continue
        
        print(f"Total elements found: {len(posts_found)}")
        
        # Extract post content
        posts_data = []
        seen_texts = set()
        
        for i, post in enumerate(posts_found[:20]):  # Limit to 20 posts
            try:
                post_text = post.text.strip()
                if post_text and len(post_text) > 20 and post_text not in seen_texts:  # Increased minimum length
                    seen_texts.add(post_text)
                    posts_data.append({
                        'index': i + 1,
                        'text': post_text,
                        'length': len(post_text),
                        'timestamp': datetime.now().isoformat()
                    })
            except Exception as e:
                print(f"Error extracting post {i}: {e}")
        
        # Create text file
        posts_text = ''
        posts_file = None
        
        if posts_data:
            posts_text = '\n\n--- POST ---\n\n'.join([f"Post {p['index']}:\n{p['text']}" for p in posts_data])
            posts_file = os.path.join(DATA_DIR, f"{company_name.replace(' ', '_')}_posts.txt")
            with open(posts_file, 'w', encoding='utf-8') as f:
                f.write(posts_text)
        
        return posts_data, posts_text, posts_file
        
    except Exception as e:
        print(f'Error getting posts for {linkedin_url}: {e}')
        return [], '', None
    finally:
        if driver:
            driver.quit()

def main():
    """Main function"""
    print("Simple LinkedIn Posts Scraper (No Login)")
    print("Targeting: core-section-container__content break-words")
    print("Using main company page instead of posts page")
    print("=" * 60)
    
    for i, row in enumerate(records, start=2):
        company_name = row.get(COMPANY_NAME_COL)
        linkedin_url = row.get(LINKEDIN_URL_COL)
        
        if not linkedin_url or linkedin_url == 'NOT FOUND':
            print(f"Skipping {company_name}: No LinkedIn URL")
            continue
        
        # Check if already processed
        existing_posts = row.get(LINKEDIN_POSTS_COL)
        if existing_posts:
            print(f"Skipping {company_name}: Already has posts")
            continue
        
        print(f"\nProcessing: {company_name}")
        
        # Get posts
        posts_data, posts_text, posts_file = get_linkedin_posts(linkedin_url, company_name)
        
        # Update spreadsheet
        if posts_text:
            worksheet.update_cell(i, linkedin_posts_col_idx, posts_text[:500])
        if posts_file:
            worksheet.update_cell(i, linkedin_posts_file_col_idx, posts_file)
        
        print(f"Extracted {len(posts_data)} posts for {company_name}")
        
        # Delay between companies
        random_delay(5, 10)
    
    print("\nScraping completed!")

if __name__ == "__main__":
    main() 