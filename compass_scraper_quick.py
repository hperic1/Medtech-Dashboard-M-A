"""
QUICK START: ISI COMPASS Scraper
Simple version - Just update the config and run!
"""

# ============= CONFIGURATION - UPDATE THESE =============
LOGIN_URL = "https://isicompass.com/login"  # Your actual login URL
DEALS_URL = "https://isicompass.com/deals"  # Your actual deals page URL
USERNAME = "your_email@example.com"          # Your ISI COMPASS email
PASSWORD = "your_password"                   # Your ISI COMPASS password

# CSS Selectors - Update if needed
USERNAME_SELECTOR = "input[type='email']"    # How to find username field
PASSWORD_SELECTOR = "input[type='password']" # How to find password field
LOGIN_BUTTON_SELECTOR = "button[type='submit']"  # Login button
TABLE_ROW_SELECTOR = "tbody tr"              # Table rows
TABLE_CELL_SELECTOR = "td"                   # Table cells
# ========================================================

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime

def quick_scrape():
    """Simple scraping function"""
    print("🚀 Starting ISI COMPASS scraper...")
    
    # Setup Chrome
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    
    try:
        # Step 1: Login
        print("\n📝 Logging in...")
        driver.get(LOGIN_URL)
        time.sleep(3)
        
        driver.find_element(By.CSS_SELECTOR, USERNAME_SELECTOR).send_keys(USERNAME)
        driver.find_element(By.CSS_SELECTOR, PASSWORD_SELECTOR).send_keys(PASSWORD)
        driver.find_element(By.CSS_SELECTOR, LOGIN_BUTTON_SELECTOR).click()
        
        time.sleep(5)  # Wait for login
        print("✅ Logged in!")
        
        # Step 2: Go to deals
        print("\n📊 Loading deals page...")
        driver.get(DEALS_URL)
        time.sleep(5)
        
        # Wait for table
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, TABLE_ROW_SELECTOR))
        )
        print("✅ Deals loaded!")
        
        # Step 3: Extract data
        print("\n🔍 Extracting deals...")
        rows = driver.find_elements(By.CSS_SELECTOR, TABLE_ROW_SELECTOR)
        
        deals_text = []
        count = 0
        
        for row in rows:
            try:
                cells = row.find_elements(By.CSS_SELECTOR, TABLE_CELL_SELECTOR)
                
                if len(cells) >= 5:
                    company = cells[0].text.strip()
                    deal_type = cells[1].text.strip()
                    technology = cells[2].text.strip()
                    investors = cells[3].text.strip()
                    amount = cells[4].text.strip()
                    date = cells[5].text.strip() if len(cells) > 5 else ""
                    
                    # Format for dashboard
                    if "Mergers and Acquisitions" in deal_type:
                        # Extract acquirer from investors column
                        acquirer = investors.split()[0] if investors else "Undisclosed"
                        
                        deal_text = f"{acquirer}—{company}\n"
                        deal_text += f"{technology}\n"
                        if date:
                            deal_text += f"Date: {date}\n"
                        deal_text += f"Value: {amount}\n"
                        
                        deals_text.append(deal_text)
                        count += 1
                        print(f"  ✓ {count}. {company}")
            
            except Exception as e:
                continue
        
        # Step 4: Save results
        print(f"\n💾 Saving {count} deals...")
        
        filename = f'compass_deals_{datetime.now().strftime("%Y%m%d_%H%M")}.txt'
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(deals_text))
        
        print(f"✅ Saved to: {filename}")
        
        # Final instructions
        print("\n" + "="*60)
        print("🎉 EXTRACTION COMPLETE!")
        print("="*60)
        print(f"\nNext steps:")
        print(f"1. Open: {filename}")
        print(f"2. Copy all text (Ctrl+A, Ctrl+C)")
        print(f"3. Open your dashboard → Data Management → Data Extraction")
        print(f"4. Paste the text")
        print(f"5. Click 'Extract Deals'")
        print(f"6. Review and add to dashboard!")
        print("="*60)
        
        input("\n\nPress Enter to close browser...")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        print("\n💡 Troubleshooting:")
        print("1. Check your LOGIN_URL and DEALS_URL")
        print("2. Verify your USERNAME and PASSWORD")
        print("3. Update CSS selectors if website structure changed")
        input("\nPress Enter to close...")
    
    finally:
        driver.quit()

if __name__ == "__main__":
    print("""
╔══════════════════════════════════════════════════════════╗
║          ISI COMPASS Automated Deal Scraper              ║
║                    Quick Start Version                   ║
╚══════════════════════════════════════════════════════════╝

Before running:
1. Install requirements: pip install selenium
2. Install ChromeDriver: brew install chromedriver (Mac)
3. Update configuration at top of this file
4. Save and run: python compass_scraper_quick.py

""")
    
    input("Press Enter to start scraping...")
    quick_scrape()
