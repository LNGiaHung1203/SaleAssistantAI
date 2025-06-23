import requests
from bs4 import BeautifulSoup

def find_linkedin_url(company_name: str) -> str:
    """Search Google for the company's LinkedIn page using Selenium, fallback to requests if needed."""
    query = f"{company_name} site:linkedin.com/company"
    google_url = f"https://www.google.com/search?q={requests.utils.quote(query)}"
    try:
        # Try Selenium first
        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.chrome.options import Options
            import time
            options = Options()
            # Do NOT use headless mode to avoid Google bot detection
            # options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            driver = webdriver.Chrome(options=options)
            driver.get(google_url)
            time.sleep(2)
            # Accept cookies if prompted
            try:
                consent = driver.find_element(By.XPATH, "//button[contains(., 'I agree') or contains(., 'Accept all') or contains(., 'Accept')]")
                consent.click()
                time.sleep(1)
            except Exception:
                pass
            # Get search results
            links = driver.find_elements(By.XPATH, "//a")
            for link in links:
                href = link.get_attribute('href')
                if href and ('linkedin.com/company/' in href or 'linkedin.com/org/' in href):
                    driver.quit()
                    return href
            driver.quit()
        except Exception as e:
            print(f"Selenium search failed: {e}. Falling back to requests.")
        # Fallback to requests
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        }
        resp = requests.get(google_url, headers=headers, timeout=10)
        soup = BeautifulSoup(resp.text, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if 'linkedin.com/company/' in href or 'linkedin.com/org/' in href:
                if href.startswith('/url?q='):
                    href = href.split('/url?q=')[1].split('&')[0]
                return href
    except Exception as e:
        print(f"Error searching for {company_name}: {e}")
    return None 