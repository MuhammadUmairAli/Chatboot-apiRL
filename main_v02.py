import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIG ===
MAX_PRODUCTS = 16
URL = "https://rubyleather.net/collections/cowhide-and-shearling-jackets-sheepskin-jackets-men/fur"
SHEET_NAME = "Designer Inspired Jackets"
EXCEL_FILE = "ruby_leather_products_updated_final.xlsx"
CHROMEDRIVER_PATH = r"E:\AnnovaSol\Rubby Leather Chat bot\chromedriver-win64\chromedriver-win64\chromedriver.exe"

# === DRIVER SETUP ===
service = Service(CHROMEDRIVER_PATH)
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--log-level=3")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(service=service, options=options)

def get_product_sizes(link, retries=2):
    for attempt in range(retries):
        try:
            driver.set_page_load_timeout(20)
            driver.get(link)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product__variant-option")))
            sizes = driver.find_elements(By.CSS_SELECTOR, "div.product__variant-option label span")
            return ", ".join([s.text.strip() for s in sizes if s.text.strip()]) or "Not listed"
        except Exception as e:
            print(f"  [SKIP] Retry {attempt + 1}/{retries} failed for sizes at {link}: {e}")
            time.sleep(1)
    return "Error"


def scroll_page():
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def scrape_products(url):
    try:
        print(f"🔄 Scraping: {url}")
        driver.get(url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.grid__item")))
        scroll_page()
        cards = driver.find_elements(By.CSS_SELECTOR, "li.grid__item")
        print(f"[INFO] Found {len(cards)} products")

        products = []
        for idx, card in enumerate(cards[:MAX_PRODUCTS]):
            try:
                # === TITLE ===
                try:
                    title = card.find_element(By.CSS_SELECTOR, "h3.card__heading").text.strip()
                except Exception:
                    try:
                        title = card.text.split("\n")[0].strip()
                    except Exception:
                        title = "Not Found"

                # === PRICE ===
                try:
                    price = card.find_element(By.CSS_SELECTOR, "span.price-item--sale").text.strip()
                except:
                    try:
                        price = card.find_element(By.CSS_SELECTOR, "span.price-item").text.strip()
                    except:
                        price = "Not listed"

                # === LINK ===
                try:
                    link = card.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
                except:
                    link = "Not Found"

                # === SIZES from detail page ===
                sizes = get_product_sizes(link) if "http" in link else "No link"

                products.append({
                    "title": title,
                    "price": price,
                    "sizes": sizes,
                    "link": link
                })

            except Exception as e:
                print(f"  [SKIP] Could not parse item {idx + 1}: {e}")

        return products

    except Exception as e:
        print(f"❌ Retry failed: {e}")
        return []


# === RUN ===
products = scrape_products(URL)
driver.quit()

# Save to Excel
if products:
    df = pd.DataFrame(products)
    try:
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        print(f"✅ Data saved under sheet: {SHEET_NAME}")
    except PermissionError:
        temp_file = "TEMP_" + EXCEL_FILE
        df.to_excel(temp_file, sheet_name=SHEET_NAME, index=False)
        print(f"⚠️ Excel file was open. Data saved to temp file: {temp_file}")
else:
    print("⚠️ No data scraped.")
