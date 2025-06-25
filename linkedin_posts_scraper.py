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
LINKEDIN_USERNAME = 'edwardwateryhung@gmail.com'
LINKEDIN_PASSWORD = 'giahung1232003'
DATA_DIR = 'company_data'

# === ANTI-DETECTION CONFIGURATION ===
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
]

# === CAPTCHA BYPASS CONFIGURATION ===
CAPTCHA_TIMEOUT = 120  # 2 minutes to solve captcha
CAPTCHA_RETRY_ATTEMPTS = 3

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
    """Setup Chrome driver with anti-detection measures"""
    chrome_options = Options()
    chrome_options.binary_location = BRAVE_PATH
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-plugins')
    chrome_options.add_argument('--disable-images')  # Faster loading
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Random user agent
    user_agent = random.choice(USER_AGENTS)
    chrome_options.add_argument(f'--user-agent={user_agent}')
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def random_delay(min_seconds=3, max_seconds=8):
    """Add random delay to prevent detection"""
    delay = random.uniform(min_seconds, max_seconds)
    print(f"Waiting {delay:.2f} seconds...")
    time.sleep(delay)

def detect_captcha(driver):
    """Detect various types of captchas on LinkedIn"""
    captcha_selectors = [
        # LinkedIn security check
        'div[data-test-id="challenge-dialog"]',
        'div.challenge-dialog',
        'div.security-verification',
        
        # Puzzle captcha
        'div[data-test-id="puzzle-captcha"]',
        'div.puzzle-captcha',
        'canvas[data-test-id="puzzle-canvas"]',
        
        # Image captcha
        'div[data-test-id="image-captcha"]',
        'div.image-captcha',
        'img[data-test-id="captcha-image"]',
        
        # General security challenges
        'div[data-test-id="challenge"]',
        'div.challenge',
        'div.security-challenge',
        
        # Phone verification
        'div[data-test-id="phone-verification"]',
        'div.phone-verification',
        
        # Email verification
        'div[data-test-id="email-verification"]',
        'div.email-verification',
        
        # Generic captcha indicators
        'iframe[src*="captcha"]',
        'div[class*="captcha"]',
        'div[class*="challenge"]',
        'div[class*="verification"]'
    ]
    
    for selector in captcha_selectors:
        try:
            element = driver.find_element(By.CSS_SELECTOR, selector)
            if element.is_displayed():
                print(f"Captcha detected with selector: {selector}")
                return True, selector
        except NoSuchElementException:
            continue
    
    # Check for captcha in page source
    page_source = driver.page_source.lower()
    captcha_indicators = [
        'captcha', 'challenge', 'verification', 'security check',
        'prove you\'re human', 'robot check', 'puzzle'
    ]
    
    for indicator in captcha_indicators:
        if indicator in page_source:
            print(f"Captcha indicator found in page: {indicator}")
            return True, "page_source"
    
    return False, None

def handle_puzzle_captcha(driver, captcha_type):
    """Handle puzzle captcha with manual intervention"""
    print(f"\n=== CAPTCHA DETECTED: {captcha_type} ===")
    print("Please solve the captcha manually in the browser.")
    print("The script will wait for you to complete it.")
    print("=" * 50)
    
    # Wait for user to solve captcha
    start_time = time.time()
    while time.time() - start_time < CAPTCHA_TIMEOUT:
        try:
            # Check if captcha is still present
            captcha_present, _ = detect_captcha(driver)
            if not captcha_present:
                print("Captcha appears to be solved!")
                random_delay(2, 4)
                return True
            
            # Check if we've been redirected to a success page
            current_url = driver.current_url
            if 'feed' in current_url or 'mynetwork' in current_url or 'checkpoint' not in current_url:
                print("Successfully passed captcha!")
                return True
            
            print("Waiting for captcha to be solved... (Press Ctrl+C to abort)")
            time.sleep(5)
            
        except KeyboardInterrupt:
            print("\nCaptcha solving aborted by user.")
            return False
        except Exception as e:
            print(f"Error checking captcha status: {e}")
            time.sleep(5)
    
    print("Captcha timeout reached. Please try again.")
    return False

def linkedin_auto_login(driver, username, password):
    """Login to LinkedIn with captcha handling"""
    try:
        driver.get('https://www.linkedin.com/login')
        random_delay(2, 4)
        
        # Wait for page to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'username'))
        )
        
        # Check for captcha before login
        captcha_present, captcha_type = detect_captcha(driver)
        if captcha_present:
            if not handle_puzzle_captcha(driver, captcha_type):
                return False
        
        # Random typing delays
        user_input = driver.find_element(By.ID, 'username')
        for char in username:
            user_input.send_keys(char)
            time.sleep(random.uniform(0.05, 0.15))
        
        random_delay(1, 2)
        
        pass_input = driver.find_element(By.ID, 'password')
        for char in password:
            pass_input.send_keys(char)
            time.sleep(random.uniform(0.05, 0.15))
        
        random_delay(1, 2)
        pass_input.submit()
        
        # Wait for login to complete
        random_delay(5, 8)
        
        # Check for captcha after login attempt
        captcha_present, captcha_type = detect_captcha(driver)
        if captcha_present:
            if not handle_puzzle_captcha(driver, captcha_type):
                return False
        
        # Check if login was successful
        if 'feed' in driver.current_url or 'mynetwork' in driver.current_url:
            print("LinkedIn login successful")
            return True
        else:
            print("LinkedIn login may have failed")
            return False
            
    except Exception as e:
        print(f"Error during LinkedIn login: {e}")
        return False

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

def extract_linkedin_posts(linkedin_url, company_name):
    """Extract LinkedIn posts with enhanced anti-detection and captcha handling"""
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
        
        # Login to LinkedIn with captcha handling
        if not linkedin_auto_login(driver, LINKEDIN_USERNAME, LINKEDIN_PASSWORD):
            return posts_data, posts_text, posts_file, posts_json_file
        
        # Navigate to posts page
        print(f"Navigating to posts page: {posts_url}")
        driver.get(posts_url)
        random_delay(5, 8)
        
        # Check for captcha on posts page
        captcha_present, captcha_type = detect_captcha(driver)
        if captcha_present:
            if not handle_puzzle_captcha(driver, captcha_type):
                return posts_data, posts_text, posts_file, posts_json_file
        
        # Scroll to load more posts
        scroll_page(driver, scroll_pause_time=3, max_scrolls=3)
        
        # Extract posts with multiple selectors
        post_selectors = [
            'div.feed-shared-update-v2__description',
            'div.update-components-text',
            'div.feed-shared-text',
            'span.break-words',
            'div.feed-shared-text__text-view'
        ]
        
        posts_found = []
        for selector in post_selectors:
            try:
                posts = driver.find_elements(By.CSS_SELECTOR, selector)
                if posts:
                    posts_found.extend(posts)
                    break
            except Exception:
                continue
        
        # Extract post content
        for i, post in enumerate(posts_found[:20]):  # Limit to 20 posts
            try:
                post_text = post.text.strip()
                if post_text and len(post_text) > 10:  # Filter out very short posts
                    post_data = {
                        'index': i + 1,
                        'text': post_text,
                        'length': len(post_text),
                        'timestamp': datetime.now().isoformat()
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
    """Main function to process all companies"""
    print("Starting LinkedIn posts scraper with captcha bypass...")
    print("Note: If captchas appear, solve them manually in the browser window.")
    print("The script will wait for you to complete them.")
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
        
        # Extract posts
        posts_data, posts_text, posts_file, posts_json_file = extract_linkedin_posts(linkedin_url, company_name)
        
        # Update spreadsheet
        if posts_text:
            worksheet.update_cell(i, linkedin_posts_col_idx, posts_text[:500])  # Preview
        if posts_file:
            worksheet.update_cell(i, linkedin_posts_file_col_idx, posts_file)
        if posts_json_file:
            worksheet.update_cell(i, linkedin_posts_json_col_idx, posts_json_file)
        
        # Random delay between companies
        random_delay(10, 20)
    
    print("\nLinkedIn posts scraping completed!")

if __name__ == "__main__":
    main() 