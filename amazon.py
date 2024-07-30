from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time
import random
from tqdm import tqdm
from datetime import datetime

def scrape_with_selenium(url):
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)

    driver.get(url)

    wait = WebDriverWait(driver, 10)
    try:
        price_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.a-price-whole')))
        price = price_element.text
    except:
        price = None

    try:
        name_element = wait.until(EC.presence_of_element_located((By.ID, 'productTitle')))
        product_name = name_element.text.strip()
    except:
        product_name = None

    driver.quit()

    return {"product": product_name, "price": price, "url": url}

output_file = "product_data.xlsx"

# Load existing data if the file exists and is valid
if os.path.exists(output_file):
    try:
        df = pd.read_excel(output_file, engine='openpyxl')
    except Exception as e:
        print(f"Error reading {output_file}: {e}")
        df = pd.DataFrame(columns=["date", "time", "product", "price", "url"])
else:
    df = pd.DataFrame(columns=["date", "time", "product", "price", "url"])

with open("urls.txt", 'r') as urllist:
    urls = urllist.read().splitlines()
    total_urls = len(urls)

    for url in tqdm(urls, total=total_urls, desc="Scraping URLs"):
        data = scrape_with_selenium(url)
        if data:
            now = datetime.now()
            data['date'] = now.date()
            data['time'] = now.time()
            new_df = pd.DataFrame([data])
            df = pd.concat([df, new_df], ignore_index=True)
        time.sleep(random.uniform(1, 5))  # Random delay between 1 and 5 seconds

# Save the DataFrame to Excel, using openpyxl engine
try:
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Data scraped successfully and saved to {output_file}")
except Exception as e:
    print(f"Error writing to {output_file}: {e}")