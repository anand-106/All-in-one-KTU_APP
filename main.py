import time
import re
import pandas as pd

import warnings

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Suppress unnecessary warnings
warnings.filterwarnings("ignore", category=UserWarning)

# ======================================
# CONFIGURATION (UPDATE THESE VALUES)
# ======================================
SEARCH_QUERY = "chennai floods"
CHROME_PROFILE_PATH = r"C:\Users\gamin\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE_NAME = "Default"
OUTPUT_FILE = "emergency_tweets.csv"
RAW_OUTPUT_FILE = "all_potential_tweets.csv"  # New file for all tweets
SCROLL_ITERATIONS = 5

EMERGENCY_KEYWORDS = [
    "help", "urgent", "stranded", "rescue", "trapped",
    "stuck", "need help", "save us", "emergency", "sos"
]

LOCATIONS = [
    "Chennai", "Adyar", "Anna Nagar", "Velachery",
    "Tambaram", "Pallavaram", "Thiruvanmiyur"
]

# ======================================
# CHROME SETUP WITH USER PROFILE
# ======================================
def init_driver():
    chrome_options = Options()
    
    # Chrome profile configuration
    chrome_options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")
    chrome_options.add_argument(f"profile-directory={CHROME_PROFILE_NAME}")
    
    # Browser settings
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    # Initialize driver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# ======================================
# IMPROVED TWEET SCRAPING FUNCTIONS
# ======================================
def scroll_page(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(SCROLL_ITERATIONS):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def scrape_tweets(driver):
    print("🚨 Starting emergency tweet scraping using your Chrome profile...")
    
    try:
        # Open Twitter search with modern parameters
        driver.get(f"https://x.com/search?q={SEARCH_QUERY}&src=spelling_expansion_revert_click&f=live")
        
        # Wait for content to load
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//article[@data-testid="tweet"]'))
        )
        
        # Scroll to load content
        scroll_page(driver)
        
        # Find tweets using modern selectors
        tweets = driver.find_elements(By.XPATH, '//article[@data-testid="tweet"]')
        print(f"✅ Found {len(tweets)} potential tweets")

        processed_data = []
        
        for tweet in tweets:
            try:
                # Improved element extraction with null checks
                text_element = WebDriverWait(tweet, 5).until(
                    EC.presence_of_element_located((By.XPATH, './/div[@data-testid="tweetText"]'))
                )
                text = text_element.text
                
                if "RT @" in text:  # Skip retweets
                    continue
                
                # User information extraction
                user_element = WebDriverWait(tweet, 5).until(
                    EC.presence_of_element_located((By.XPATH, './/div[@data-testid="User-Name"]'))
                )
                user = user_element.text.split('\n')[0]
                
                # Timestamp and URL
                time_element = WebDriverWait(tweet, 5).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'time'))
                )
                timestamp = time_element.get_attribute('datetime')
                tweet_url = time_element.find_element(By.XPATH, './..').get_attribute('href')

                processed_data.append({
                    "text": text,
                    "user": user,
                    "timestamp": timestamp,
                    "url": tweet_url
                })
                
            except Exception as e:
                continue
        
        return processed_data

    except Exception as e:
        print(f"⚠️ Error during scraping: {str(e)}")
        return []

# ======================================
# IMPROVED DATA PROCESSING
# ======================================
def process_data(raw_data):
    if not raw_data:
        return pd.DataFrame(), pd.DataFrame()
    
    df = pd.DataFrame(raw_data)
    
    # Add analysis columns
    emergency_pattern = re.compile(r'\b(?:' + '|'.join(EMERGENCY_KEYWORDS) + r')\b', flags=re.IGNORECASE)
    df['emergency'] = df['text'].str.contains(emergency_pattern)
    
    location_pattern = re.compile(r'\b(?:' + '|'.join(LOCATIONS) + r')\b', flags=re.IGNORECASE)
    df['locations'] = df['text'].str.findall(location_pattern)
    
    # Create filtered dataset
    filtered_df = df[(df['emergency']) & (df['locations'].str.len() > 0)]
    
    return df, filtered_df  # Return both datasets

# Modified main execution
if __name__ == "__main__":
    driver = init_driver()
    
    try:
        raw_data = scrape_tweets(driver)
        
        if raw_data:
            all_tweets_df, emergency_df = process_data(raw_data)
            
            # Save all potential tweets
            all_tweets_df.to_csv(RAW_OUTPUT_FILE, index=False)
            print(f"\n💾 Saved all potential tweets to {RAW_OUTPUT_FILE}")
            
            if not emergency_df.empty:
                print(f"\n🚨 Found {len(emergency_df)} emergency tweets:")
                print(emergency_df[['text', 'locations', 'timestamp']].head())
                emergency_df.to_csv(OUTPUT_FILE, index=False)
                print(f"\n💾 Saved emergency tweets to {OUTPUT_FILE}")
            else:
                print("\n⚠️ No emergency tweets found matching criteria")
        else:
            print("\n⚠️ No tweets found for the search query")
            
    finally:
        driver.quit()
