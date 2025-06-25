# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import random
import re
import os
import json
from datetime import datetime, timedelta
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# === CONFIGURATION ===
BRAVE_PATH = r'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe'
DATA_DIR = 'company_data'
EXCEL_FILE = 'Company list.xlsx'  # Excel file with company names
RESULTS_FILE = 'company_results.json'  # Output file with all results

os.makedirs(DATA_DIR, exist_ok=True)

def setup_driver():
    chrome_options = Options()
    chrome_options.binary_location = BRAVE_PATH
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-plugins')
    chrome_options.add_argument('--disable-images')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0'
    ]
    user_agent = random.choice(user_agents)
    chrome_options.add_argument(f'--user-agent={user_agent}')
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def random_delay(min_seconds=2, max_seconds=6):
    delay = random.uniform(min_seconds, max_seconds)
    print(f"Waiting {delay:.2f} seconds...")
    time.sleep(delay)

def close_linkedin_popup(driver):
    """Close LinkedIn login popup if it appears"""
    try:
        # Wait a bit for popup to appear
        random_delay(2, 4)
        
        # Primary selector - the specific LinkedIn close button you provided
        primary_selector = 'button.modal__dismiss.btn-tertiary.h-\\[40px\\].w-\\[40px\\].p-0.rounded-full.indent-0.sign-in-modal__dismiss.absolute.right-0.cursor-pointer.m-\\[20px\\]'
        
        try:
            close_button = driver.find_element(By.CSS_SELECTOR, primary_selector)
            if close_button.is_displayed():
                print("Found and closing LinkedIn popup with primary selector")
                close_button.click()
                random_delay(1, 2)
                return True
        except:
            pass
        
        # Alternative selectors for the same button (simplified versions)
        alternative_selectors = [
            'button.modal__dismiss',
            'button.sign-in-modal__dismiss',
            'button[class*="modal__dismiss"]',
            'button[class*="sign-in-modal__dismiss"]',
            '.modal__dismiss',
            '.sign-in-modal__dismiss'
        ]
        
        for selector in alternative_selectors:
            try:
                close_button = driver.find_element(By.CSS_SELECTOR, selector)
                if close_button.is_displayed():
                    print(f"Found and closing popup with selector: {selector}")
                    close_button.click()
                    random_delay(1, 2)
                    return True
            except:
                continue
        
        # Fallback selectors for other types of popups
        fallback_selectors = [
            'button[aria-label="Dismiss"]',
            'button[aria-label="Close"]',
            'button[data-test-id="close-button"]',
            'button.close',
            'button[class*="close"]',
            'button[class*="dismiss"]',
            'div[class*="close"] button',
            'div[class*="dismiss"] button',
            'button:contains("×")',
            'button:contains("Close")',
            'button:contains("Dismiss")',
            'button:contains("Skip")',
            'button:contains("Not now")',
            'button:contains("Maybe later")'
        ]
        
        for selector in fallback_selectors:
            try:
                close_button = driver.find_element(By.CSS_SELECTOR, selector)
                if close_button.is_displayed():
                    print(f"Found and closing popup with fallback selector: {selector}")
                    close_button.click()
                    random_delay(1, 2)
                    return True
            except:
                continue
        
        # Try clicking outside the popup
        try:
            driver.find_element(By.TAG_NAME, "body").click()
            print("Clicked outside popup to close it")
            random_delay(1, 2)
            return True
        except:
            pass
        
        # Try pressing Escape key
        try:
            from selenium.webdriver.common.keys import Keys
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            print("Pressed Escape to close popup")
            random_delay(1, 2)
            return True
        except:
            pass
            
    except Exception as e:
        print(f"Error closing popup: {e}")
    
    return False

def find_linkedin_url(company_name, driver):
    print(f"Searching for LinkedIn URL: {company_name}")
    query = f"{company_name} site:linkedin.com/company"
    google_url = f"https://www.google.com/search?q={requests.utils.quote(query)}"
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            driver.get(google_url)
            random_delay(2, 4)
            
            # Accept Google cookies if prompted
            try:
                consent_selectors = [
                    "button[aria-label*='Accept']",
                    "button[aria-label*='Agree']",
                    "button:contains('Accept all')",
                    "button:contains('I agree')"
                ]
                for selector in consent_selectors:
                    try:
                        consent = driver.find_element(By.CSS_SELECTOR, selector)
                        consent.click()
                        random_delay(1, 2)
                        break
                    except:
                        continue
            except Exception:
                pass

            # Detect and close Google's hidden popup with id='close' and class 'TvD9Pc-Bz112c'
            try:
                close_popup = driver.find_element(By.CSS_SELECTOR, '#close.TvD9Pc-Bz112c')
                if close_popup.is_displayed():
                    print("Found and closing Google hidden popup with id='close' and class 'TvD9Pc-Bz112c'")
                    close_popup.click()
                    random_delay(1, 2)
            except Exception:
                pass

            # CAPTCHA detection and pause for manual solving
            def is_captcha_present():
                try:
                    if driver.find_elements(By.ID, 'captcha-form'):
                        return True
                    if driver.find_elements(By.XPATH, "//*[contains(text(), 'not a robot') or contains(text(), 'unusual traffic')]"):
                        return True
                    if driver.find_elements(By.CSS_SELECTOR, 'div#g-recaptcha, div.recaptcha'):
                        return True
                    return False
                except:
                    return False
            if is_captcha_present():
                print(f"[!!!] Google CAPTCHA detected for {company_name}. Please solve the CAPTCHA in the opened browser window.")
                input("Press Enter here after you have solved the CAPTCHA in the browser...")
                # Optionally, re-check and wait until CAPTCHA is gone
                while is_captcha_present():
                    print("CAPTCHA still detected. Please solve it in the browser.")
                    input("Press Enter again after solving...")

            links = driver.find_elements(By.XPATH, "//a")
            for link in links:
                href = link.get_attribute('href')
                if href and ('linkedin.com/company/' in href or 'linkedin.com/org/' in href):
                    if href.startswith('/url?q='):
                        href = href.split('/url?q=')[1].split('&')[0]
                    print(f"Found LinkedIn URL: {href}")
                    return href
        except Exception as e:
            print(f"Error searching for {company_name}: {e}")
    print(f"[!] Skipping {company_name} after {max_retries} failed attempts due to CAPTCHA.")
    return None

def check_hiring_status(linkedin_url, company_name, driver):
    if not linkedin_url:
        return {'is_hiring': 'Unknown', 'num_jobs': 0, 'jobs_url': None, 'source': 'no_url'}
    print(f"Checking hiring status for: {company_name}")
    base_url = linkedin_url.split('?')[0].rstrip('/')
    jobs_url = base_url + '/jobs/'
    try:
        driver.get(jobs_url)
        random_delay(3, 6)
        
        # Close any popup that appears
        close_linkedin_popup(driver)
        
        current_url = driver.current_url
        print(f"Current URL after accessing jobs page: {current_url}")
        
        # Check if we were redirected away from jobs page
        if 'login' in current_url or 'signup' in current_url:
            print("Jobs page requires login, trying main company page...")
            return check_hiring_from_main_page(driver, linkedin_url, company_name)
        
        # Check if we're still on the jobs page or were redirected
        if '/jobs' not in current_url and jobs_url not in current_url:
            print("Jobs page redirected away - likely no open positions")
            return {
                'is_hiring': 'No open positions',
                'num_jobs': 0,
                'jobs_url': jobs_url,
                'source': 'jobs_redirect',
                'redirected_to': current_url
            }
        
        # If we're still on jobs page, company is hiring
        print("Jobs page accessible - company is hiring")
        
        # Primary selector - the specific jobs section you identified
        primary_jobs_selectors = [
            'div.core-section-container.my-3[data-test-id="jobs-at"]',
            'div[data-test-id="jobs-at"]',
            'div.core-section-container.my-3',
            'div[class*="core-section-container"][class*="my-3"]'
        ]
        
        # Check for the specific jobs section first
        for selector in primary_jobs_selectors:
            try:
                jobs_section = driver.find_element(By.CSS_SELECTOR, selector)
                if jobs_section.is_displayed():
                    print(f"Found jobs section with selector: {selector}")
                    
                    # Look for job count within this section
                    job_count_selectors = [
                        'span[class*="job-count"]',
                        'span[class*="opening"]',
                        'div[class*="job-count"]',
                        'div[class*="opening"]',
                        'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                        'span', 'div'
                    ]
                    
                    for count_selector in job_count_selectors:
                        try:
                            elements = jobs_section.find_elements(By.CSS_SELECTOR, count_selector)
                            for element in elements:
                                text = element.text.strip()
                                if text:
                                    print(f"Found job indicator in jobs section: {text}")
                                    # Extract numbers from text
                                    numbers = re.findall(r'\d+', text)
                                    if numbers:
                                        num_jobs = int(numbers[0])
                                        return {
                                            'is_hiring': 'Yes',
                                            'num_jobs': num_jobs,
                                            'jobs_url': jobs_url,
                                            'source': 'jobs_section',
                                            'detected_text': text,
                                            'section_selector': selector
                                        }
                        except Exception as e:
                            continue
                    
                    # If no specific count found, check if section has any job-related content
                    section_text = jobs_section.text.lower()
                    if 'job' in section_text or 'opening' in section_text or 'position' in section_text:
                        return {
                            'is_hiring': 'Yes (jobs section found)',
                            'num_jobs': 0,
                            'jobs_url': jobs_url,
                            'source': 'jobs_section_content',
                            'section_selector': selector
                        }
            except Exception as e:
                print(f"Error with primary jobs selector {selector}: {e}")
                continue
        
        # Enhanced job indicators with better detection (fallback)
        job_indicators = [
            'h4.org-jobs-job-search-form-module__headline',
            'span[data-test-id="job-count"]',
            'div[class*="job-count"]',
            'span[class*="job"]',
            'div[class*="hiring"]',
            'div[class*="career"]',
            'span[class*="opening"]',
            'div[class*="opening"]',
            'h1[class*="job"]',
            'h2[class*="job"]',
            'h3[class*="job"]',
            'h4[class*="job"]',
            'div[data-test-id*="job"]',
            'span[data-test-id*="job"]'
        ]
        
        for selector in job_indicators:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    text = element.text.strip()
                    if text:
                        print(f"Found job indicator: {text}")
                        # Extract numbers from text
                        numbers = re.findall(r'\d+', text)
                        if numbers:
                            num_jobs = int(numbers[0])
                            return {
                                'is_hiring': 'Yes',
                                'num_jobs': num_jobs,
                                'jobs_url': jobs_url,
                                'source': 'jobs_page',
                                'detected_text': text
                            }
            except Exception as e:
                print(f"Error with selector {selector}: {e}")
                continue
        
        # Check page source for job-related keywords
        page_text = driver.page_source.lower()
        job_keywords = ['job', 'hiring', 'career', 'opening', 'position', 'vacancy']
        job_count = 0
        
        for keyword in job_keywords:
            if keyword in page_text:
                job_count += page_text.count(keyword)
        
        if job_count > 5:  # If many job-related keywords found
            return {
                'is_hiring': 'Yes (job keywords found)',
                'num_jobs': 0,
                'jobs_url': jobs_url,
                'source': 'jobs_page_keywords',
                'keyword_count': job_count
            }
        
        # If we're on jobs page but no specific job info found, still hiring
        return {
            'is_hiring': 'Yes (jobs page accessible)',
            'num_jobs': 0,
            'jobs_url': jobs_url,
            'source': 'jobs_page_accessible'
        }
    except Exception as e:
        print(f"Error checking hiring status for {company_name}: {e}")
        return {'is_hiring': 'Unknown', 'num_jobs': 0, 'jobs_url': jobs_url, 'source': 'error', 'error': str(e)}

def check_hiring_from_main_page(driver, linkedin_url, company_name):
    """Check hiring status from main company page"""
    try:
        driver.get(linkedin_url)
        random_delay(3, 6)
        
        # Close any popup that appears
        close_linkedin_popup(driver)
        
        # Try to access jobs page from main page
        jobs_url = linkedin_url + '/jobs/'
        print(f"Trying to access jobs page from main page: {jobs_url}")
        
        driver.get(jobs_url)
        random_delay(3, 6)
        close_linkedin_popup(driver)
        
        current_url = driver.current_url
        print(f"Current URL after accessing jobs from main page: {current_url}")
        
        # Check if jobs page is accessible
        if '/jobs' in current_url and jobs_url in current_url:
            print("Jobs page accessible from main page - company is hiring")
            return {
                'is_hiring': 'Yes (jobs page accessible)',
                'num_jobs': 0,
                'jobs_url': jobs_url,
                'source': 'main_page_jobs_accessible'
            }
        else:
            print("Jobs page not accessible - no open positions")
            return {
                'is_hiring': 'No open positions',
                'num_jobs': 0,
                'jobs_url': jobs_url,
                'source': 'main_page_jobs_redirect',
                'redirected_to': current_url
            }
        
    except Exception as e:
        print(f"Error checking main page: {e}")
        return {'is_hiring': 'Error', 'num_jobs': 0, 'jobs_url': linkedin_url + '/jobs/', 'source': 'error'}

def get_linkedin_posts(linkedin_url, company_name, driver):
    if not linkedin_url:
        return []
    print(f"Getting posts for: {company_name}")
    try:
        driver.get(linkedin_url)
        random_delay(3, 6)
        
        # Close any popup that appears
        close_linkedin_popup(driver)
        
        current_url = driver.current_url
        if 'login' in current_url or 'signup' in current_url:
            print("Company page requires login, cannot get posts")
            return []
        
        # Find the updates list (ul.updates__list)
        try:
            updates_list = driver.find_element(By.CSS_SELECTOR, 'ul.updates__list')
        except Exception as e:
            print("Could not find updates__list: ", e)
            return []
        
        # Smart scrolling to load more posts
        print("Scrolling to load more posts...")
        max_scrolls = 8
        last_post_count = 0
        no_new_posts_count = 0
        for scroll_attempt in range(max_scrolls):
            driver.execute_script("arguments[0].scrollIntoView(false);", updates_list)
            random_delay(2, 4)
            li_posts = updates_list.find_elements(By.CSS_SELECTOR, 'li.mb-1')
            current_post_count = len(li_posts)
            print(f"Scroll {scroll_attempt + 1}: Found {current_post_count} posts")
            if current_post_count > last_post_count:
                last_post_count = current_post_count
                no_new_posts_count = 0
            else:
                no_new_posts_count += 1
            if no_new_posts_count >= 2 or current_post_count >= 15:
                break
        
        # Extract posts from li.mb-1
        posts_data = []
        seen_texts = set()
        three_months_ago = datetime.now() - timedelta(days=90)
        li_posts = updates_list.find_elements(By.CSS_SELECTOR, 'li.mb-1')
        print(f"Processing {len(li_posts)} posts from updates__list...")
        for i, li in enumerate(li_posts):
            try:
                # Extract post content
                post_text = ""
                content_selectors = [
                    'p[data-test-id="main-feed-activity-card__commentary"]',
                    'div.attributed-text-segment-list__content',
                    'p.attributed-text-segment-list__content',
                    'div[class*="attributed-text-segment-list__content"]'
                ]
                for selector in content_selectors:
                    try:
                        content_element = li.find_element(By.CSS_SELECTOR, selector)
                        post_text = content_element.text.strip()
                        if post_text:
                            break
                    except:
                        continue
                if not post_text or post_text in seen_texts or len(post_text) < 20:
                    continue
                seen_texts.add(post_text)
                # Extract date
                date_text = "Unknown"
                try:
                    date_element = li.find_element(By.CSS_SELECTOR, 'time')
                    date_text = date_element.text.strip()
                except:
                    pass
                # Extract engagement metrics
                reactions_count = 0
                comments_count = 0
                try:
                    reactions_element = li.find_element(By.CSS_SELECTOR, 'span[data-test-id="social-actions__reaction-count"]')
                    reactions_text = reactions_element.text.strip()
                    if reactions_text.isdigit():
                        reactions_count = int(reactions_text)
                except:
                    pass
                try:
                    comments_element = li.find_element(By.CSS_SELECTOR, 'a[data-test-id="social-actions__comments"]')
                    comments_text = comments_element.text.strip()
                    comments_match = re.search(r'(\d+)', comments_text)
                    if comments_match:
                        comments_count = int(comments_match.group(1))
                except:
                    pass
                # Check if post is recent (within 3 months)
                is_recent = True
                if date_text != "Unknown":
                    if "mo" in date_text.lower():
                        months = int(re.search(r'(\d+)', date_text).group(1))
                        is_recent = months <= 3
                    elif "d" in date_text.lower() or "w" in date_text.lower():
                        is_recent = True
                if not is_recent:
                    continue
                post_data = {
                    'index': len(posts_data) + 1,
                    'text': post_text,
                    'length': len(post_text),
                    'timestamp': datetime.now().isoformat(),
                    'post_date': date_text,
                    'reactions_count': reactions_count,
                    'comments_count': comments_count,
                    'is_recent': is_recent,
                    'selector_used': 'updates__list > li.mb-1'
                }
                posts_data.append(post_data)
                if len(posts_data) >= 10:
                    break
            except Exception as e:
                print(f"Error extracting post {i}: {e}")
        print(f"Extracted {len(posts_data)} posts for {company_name} (limited to 10 recent posts)")
        return posts_data
    except Exception as e:
        print(f"Error getting posts for {linkedin_url}: {e}")
        return []

def save_company_data(company_name, linkedin_url, hiring_data, posts_data):
    """Save all company data to local files with enhanced JSON structure"""
    # Create company directory
    company_dir = os.path.join(DATA_DIR, company_name.replace(' ', '_').replace('/', '_'))
    os.makedirs(company_dir, exist_ok=True)
    
    # Save posts to text file
    posts_file = os.path.join(company_dir, f"{company_name.replace(' ', '_')}_posts.txt")
    if posts_data:
        posts_text = '\n\n--- POST ---\n\n'.join([f"Post {p['index']}:\n{p['text']}" for p in posts_data])
        with open(posts_file, 'w', encoding='utf-8') as f:
            f.write(posts_text)
    
    # Enhanced company data structure
    company_data = {
        'company_name': company_name,
        'linkedin_url': linkedin_url,
        'hiring_status': {
            'is_hiring': hiring_data.get('is_hiring', 'Unknown'),
            'num_jobs': hiring_data.get('num_jobs', 0),
            'jobs_url': hiring_data.get('jobs_url'),
            'source': hiring_data.get('source', 'unknown'),
            'detected_text': hiring_data.get('detected_text'),
            'keyword_count': hiring_data.get('keyword_count'),
            'job_links': hiring_data.get('job_links', [])
        },
        'posts': {
            'total_posts': len(posts_data),
            'posts_list': posts_data,
            'posts_file': posts_file if posts_data else None
        },
        'metadata': {
            'timestamp': datetime.now().isoformat(),
            'processing_date': datetime.now().strftime('%Y-%m-%d'),
            'processing_time': datetime.now().strftime('%H:%M:%S')
        }
    }
    
    # Save comprehensive JSON file
    json_file = os.path.join(company_dir, f"{company_name.replace(' ', '_')}_complete_data.json")
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(company_data, f, indent=2, ensure_ascii=False)
    
    # Also save a simplified version
    simplified_data = {
        'company_name': company_name,
        'linkedin_url': linkedin_url,
        'is_hiring': hiring_data.get('is_hiring', 'Unknown'),
        'num_jobs': hiring_data.get('num_jobs', 0),
        'total_posts': len(posts_data),
        'posts_preview': [p['text'][:100] + '...' if len(p['text']) > 100 else p['text'] for p in posts_data[:5]],
        'timestamp': datetime.now().isoformat()
    }
    
    simplified_file = os.path.join(company_dir, f"{company_name.replace(' ', '_')}_summary.json")
    with open(simplified_file, 'w', encoding='utf-8') as f:
        json.dump(simplified_data, f, indent=2, ensure_ascii=False)
    
    return company_data

def load_companies_from_excel():
    """Load company names from Excel file"""
    try:
        if os.path.exists(EXCEL_FILE):
            # Try to read the Excel file
            df = pd.read_excel(EXCEL_FILE)
            print(f"Excel file columns: {df.columns.tolist()}")
            
            # Look for company name column
            company_columns = ['company_name', 'company', 'name', 'Company Name', 'Company']
            company_col = None
            
            for col in company_columns:
                if col in df.columns:
                    company_col = col
                    break
            
            if company_col is None:
                # If no standard column name found, use the first column
                company_col = df.columns[0]
                print(f"Using first column as company names: {company_col}")
            
            companies = df[company_col].dropna().astype(str).tolist()
            print(f"Loaded {len(companies)} companies from Excel file")
            return companies
        else:
            print(f"Excel file {EXCEL_FILE} not found. Creating sample companies list.")
            sample_companies = [
                "Microsoft",
                "Google",
                "Apple",
                "Amazon",
                "Meta"
            ]
            return sample_companies
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        print("Using sample companies list instead.")
        sample_companies = [
            "Microsoft",
            "Google",
            "Apple",
            "Amazon",
            "Meta"
        ]
        return sample_companies

def process_company(args):
    i, company_name, total, driver = args
    print(f"\n{'='*50}")
    print(f"Now searching company {i}/{total}: {company_name}")
    print(f"Processing {i}/{total}: {company_name}")
    print(f"{'='*50}")
    # Check if company already has posts
    company_dir = os.path.join(DATA_DIR, company_name.replace(' ', '_').replace('/', '_'))
    json_file = os.path.join(company_dir, f"{company_name.replace(' ', '_')}_complete_data.json")
    if os.path.exists(json_file):
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                posts = data.get('posts', {}).get('posts_list', [])
                if posts:
                    print(f"[SKIP] {company_name} already has {len(posts)} posts. Skipping.")
                    return {'company_name': company_name, 'skipped': True}
        except Exception as e:
            print(f"[WARN] Could not read existing data for {company_name}: {e}")
    try:
        linkedin_url = find_linkedin_url(company_name, driver)
        random_delay(3, 6)
        hiring_data = check_hiring_status(linkedin_url, company_name, driver) if linkedin_url else {'is_hiring': 'Unknown', 'num_jobs': 0, 'jobs_url': None}
        random_delay(3, 6)
        posts_data = get_linkedin_posts(linkedin_url, company_name, driver) if linkedin_url else []
        random_delay(3, 6)
        company_data = save_company_data(company_name, linkedin_url, hiring_data, posts_data)
        print(f"\n📊 Summary for {company_name}:")
        print(f"   LinkedIn URL: {linkedin_url or 'Not found'}")
        print(f"   Hiring Status: {hiring_data.get('is_hiring', 'Unknown')}")
        print(f"   Number of Jobs: {hiring_data.get('num_jobs', 0)}")
        print(f"   Posts Found: {len(posts_data)}")
        return company_data
    except Exception as e:
        print(f"Error processing {company_name}: {e}")
        return {
            'company_name': company_name,
            'linkedin_url': None,
            'hiring_status': {'is_hiring': 'Error', 'num_jobs': 0},
            'posts': [],
            'error': str(e),
            'timestamp': datetime.now().isoformat()
        }

def main():
    print("LinkedIn Company Analyzer (No Login Required)")
    print("=" * 60)
    companies = load_companies_from_excel()
    print(f"Loaded {len(companies)} companies from {EXCEL_FILE}")
    all_results = []
    total = len(companies)
    driver = setup_driver()
    try:
        with ThreadPoolExecutor(max_workers=1) as executor:  # Only one thread
            futures = [executor.submit(process_company, (i+1, company_name, total, driver)) for i, company_name in enumerate(companies)]
            for future in as_completed(futures):
                result = future.result()
                all_results.append(result)
        with open(RESULTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(all_results, f, indent=2, ensure_ascii=False)
        print(f"\n{'='*60}")
        print("ANALYSIS COMPLETED!")
        print(f"{'='*60}")
        print(f"Total companies processed: {len(companies)}")
        print(f"Results saved to: {RESULTS_FILE}")
        print(f"Individual company data saved to: {DATA_DIR}/")
        linkedin_found = sum(1 for r in all_results if r.get('linkedin_url'))
        hiring_companies = sum(1 for r in all_results if r.get('hiring_status', {}).get('is_hiring') in ['Yes', 'Likely'])
        total_posts = sum(len(r.get('posts', {}).get('posts_list', [])) for r in all_results)
        print(f"\n📈 Summary Statistics:")
        print(f"   LinkedIn URLs found: {linkedin_found}/{len(companies)}")
        print(f"   Companies hiring: {hiring_companies}")
        print(f"   Total posts collected: {total_posts}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main() 