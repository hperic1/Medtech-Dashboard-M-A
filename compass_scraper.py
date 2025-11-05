"""
ISI COMPASS Automated Deal Scraper
Logs into ISI COMPASS, extracts deal data, and prepares it for the dashboard
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import os
from datetime import datetime

class CompassScraper:
    def __init__(self, username, password):
        """Initialize the scraper with login credentials"""
        self.username = username
        self.password = password
        self.driver = None
        
    def setup_driver(self, headless=True):
        """Setup Chrome driver with options"""
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless")  # Run without opening browser
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        
        # Optional: Set user agent to avoid detection
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(10)
        
    def login(self, login_url):
        """Log into ISI COMPASS"""
        print("Logging in to ISI COMPASS...")
        self.driver.get(login_url)
        
        # Wait for login form to load
        wait = WebDriverWait(self.driver, 20)
        
        # Find and fill login fields (adjust selectors based on actual site)
        try:
            # Try common login field selectors
            username_field = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[name='username'], input[name='email']"))
            )
            password_field = self.driver.find_element(By.CSS_SELECTOR, "input[type='password'], input[name='password']")
            
            username_field.send_keys(self.username)
            password_field.send_keys(self.password)
            
            # Find and click login button
            login_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
            login_button.click()
            
            # Wait for login to complete
            time.sleep(3)
            print("✓ Logged in successfully")
            
        except Exception as e:
            print(f"❌ Login failed: {str(e)}")
            raise
    
    def navigate_to_deals(self, deals_url):
        """Navigate to the Deals page"""
        print("Navigating to Deals page...")
        self.driver.get(deals_url)
        
        # Wait for the deals table to load
        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table, [role='table']")))
        
        # Optional: Click "Mergers and Acquisitions" filter if needed
        try:
            ma_filter = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Mergers and Acquisitions')]")
            if ma_filter:
                ma_filter.click()
                time.sleep(2)
        except:
            pass  # Filter might already be applied
        
        print("✓ Deals page loaded")
    
    def extract_deals_from_table(self):
        """Extract deal data from the table"""
        print("Extracting deals...")
        
        deals = []
        
        try:
            # Wait for table rows to load
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr, [role='row']")))
            
            # Find all table rows (adjust selector based on actual structure)
            rows = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr, [role='row']")
            
            for row in rows:
                try:
                    # Extract cells (adjust based on actual table structure)
                    cells = row.find_elements(By.CSS_SELECTOR, "td, [role='cell']")
                    
                    if len(cells) >= 5:  # Ensure we have enough columns
                        deal = {
                            'company': cells[0].text.strip(),
                            'type': cells[1].text.strip(),
                            'technology': cells[2].text.strip(),
                            'investors_details': cells[3].text.strip(),
                            'amount': cells[4].text.strip(),
                            'date': cells[5].text.strip() if len(cells) > 5 else ''
                        }
                        
                        # Only add if company name exists
                        if deal['company']:
                            deals.append(deal)
                            print(f"  ✓ Extracted: {deal['company']}")
                
                except Exception as e:
                    print(f"  ⚠ Skipped row: {str(e)}")
                    continue
            
            print(f"✓ Extracted {len(deals)} deals")
            return deals
            
        except Exception as e:
            print(f"❌ Error extracting deals: {str(e)}")
            return []
    
    def scroll_and_load_all(self):
        """Scroll page to load all deals (if lazy loading)"""
        print("Loading all deals...")
        
        last_height = self.driver.execute_script("return document.body.scrollHeight")
        
        while True:
            # Scroll down
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            # Check if more content loaded
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
        
        print("✓ All deals loaded")
    
    def save_to_csv(self, deals, filename='compass_deals.csv'):
        """Save deals to CSV file"""
        if not deals:
            print("⚠ No deals to save")
            return None
        
        df = pd.DataFrame(deals)
        df.to_csv(filename, index=False)
        print(f"✓ Saved {len(deals)} deals to {filename}")
        return filename
    
    def format_for_dashboard(self, deals):
        """Format deals for the dashboard's data extraction format"""
        formatted_text = []
        
        for deal in deals:
            # Format each deal like the text parser expects
            company = deal['company']
            acquirer = self._extract_acquirer(deal['investors_details'])
            amount = deal['amount']
            date = deal['date']
            technology = deal['technology']
            
            deal_text = f"{acquirer}—{company}\n"
            deal_text += f"{technology}\n"
            deal_text += f"Date of Announcement: {date}\n"
            deal_text += f"Value: {amount}\n"
            
            formatted_text.append(deal_text)
        
        return "\n\n".join(formatted_text)
    
    def _extract_acquirer(self, investors_details):
        """Extract acquirer name from investors/deal details"""
        # Look for patterns like "Company acquired" or "Acquired by Company"
        if 'acquired' in investors_details.lower():
            # Extract company name before "acquired"
            parts = investors_details.split('acquired')
            if parts[0].strip():
                return parts[0].strip()
        
        # Otherwise return the first company name mentioned
        words = investors_details.split()
        if words:
            return words[0]
        
        return "Undisclosed"
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()
            print("✓ Browser closed")

def main():
    """Main function to run the scraper"""
    
    # Configuration
    LOGIN_URL = "https://isicompass.com/login"  # Replace with actual login URL
    DEALS_URL = "https://isicompass.com/deals"  # Replace with actual deals URL
    
    # Get credentials from environment variables (more secure)
    USERNAME = os.getenv('COMPASS_USERNAME', 'your_email@example.com')
    PASSWORD = os.getenv('COMPASS_PASSWORD', 'your_password')
    
    # Initialize scraper
    scraper = CompassScraper(USERNAME, PASSWORD)
    
    try:
        # Setup browser
        scraper.setup_driver(headless=False)  # Set to True for background operation
        
        # Login
        scraper.login(LOGIN_URL)
        
        # Navigate to deals
        scraper.navigate_to_deals(DEALS_URL)
        
        # Load all deals (if lazy loading)
        scraper.scroll_and_load_all()
        
        # Extract deals
        deals = scraper.extract_deals_from_table()
        
        # Save to CSV
        csv_file = scraper.save_to_csv(deals, f'compass_deals_{datetime.now().strftime("%Y%m%d")}.csv')
        
        # Format for dashboard
        formatted_text = scraper.format_for_dashboard(deals)
        
        # Save formatted text for easy copy-paste
        with open('compass_deals_formatted.txt', 'w') as f:
            f.write(formatted_text)
        
        print("\n" + "="*50)
        print("✅ EXTRACTION COMPLETE!")
        print("="*50)
        print(f"CSV file: {csv_file}")
        print(f"Formatted text: compass_deals_formatted.txt")
        print("\nYou can now:")
        print("1. Open compass_deals_formatted.txt")
        print("2. Copy the text")
        print("3. Paste into your dashboard's Data Extraction tab")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
    
    finally:
        # Always close the browser
        scraper.close()

if __name__ == "__main__":
    main()
