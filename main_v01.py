from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from docx import Document
import pandas as pd
import re
import time
from bs4 import BeautifulSoup

# Setup Chrome
service = Service(r"E:\AnnovaSol\Rubby Leather Chat bot\chromedriver-win64\chromedriver-win64\chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--log-level=3")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(service=service, options=options)

# Load URLs from DOCX
doc = Document("Ruby Leather.docx")
url_pattern = re.compile(r"https?://\S+")
urls = [re.search(url_pattern, para.text).group() for para in doc.paragraphs if re.search(url_pattern, para.text)]

# Sheet name mapping
sheet_name_map = {
    "biker": "Biker Jackets",
    "designer": "Designer Inspired Jackets",
    "fur": "Fur and Shearling",
    "bomber": "Bomber Jackets",
    "celebrity": "Seen on Celebrities",
    "ladies-jacket": "Women Leather Jackets",
    "backpack": "Backpack",
    "wallet": "Wallets",
    "shoulder-bag": "Handbags",
    "duffel": "Travel Bags",
    "laptop-bag": "Laptop Bags",
    "toiletry": "Makeup and Toiletry",
    "accessories": "Accessories"
}

# Master dictionary to collect data per sheet
category_data = {}

def get_sizes_from_detail(product_url):
    try:
        driver.get(product_url)
        time.sleep(2)
        size_elements = driver.find_elements(By.CSS_SELECTOR, "div.product__variant-option label span")
        sizes = [el.text.strip() for el in size_elements if el.text.strip()]
        return ", ".join(sizes) if sizes else "Not listed"
    except:
        return "Error"

# Scraper
def scrape_collection(url):
    print(f"🔗 Scraping: {url}")
    driver.get(url)
    time.sleep(3)

    products = []
    soup = BeautifulSoup(driver.page_source, "html.parser")
    items = soup.select("li.grid__item")

    for item in items:
        try:
            title = item.select_one("h3.card__heading")
            price = item.select_one("span.price-item--sale") or item.select_one("span.price-item")
            link_tag = item.select_one("a")

            title_text = title.get_text(strip=True) if title else "N/A"
            price_text = price.get_text(strip=True) if price else "Not listed"
            link = "https://rubyleather.net" + link_tag["href"] if link_tag and "href" in link_tag.attrs else "Link not found"

            sizes = get_sizes_from_detail(link) if "http" in link else "No link"

            products.append({
                "title": title_text,
                "price": price_text,
                "sizes": sizes,
                "link": link
            })
        except Exception as e:
            print("  [SKIP] Error parsing item:", e)

    return products

# Loop through each URL
for url in urls:
    try:
        sheet_key = next((key for key in sheet_name_map if key in url.lower()), None)
        sheet_name = sheet_name_map.get(sheet_key, "Uncategorized")

        scraped = scrape_collection(url)

        if sheet_name not in category_data:
            category_data[sheet_name] = []

        category_data[sheet_name].extend(scraped)
        print(f"[DEBUG] Added {len(scraped)} items to '{sheet_name}'")

    except Exception as e:
        print(f"⚠️ Failed on {url}: {e}")

driver.quit()

# Save to Excel
output_file = "ruby_leather_products_updated_final.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for category, products in category_data.items():
        df = pd.DataFrame(products)
        if not df.empty:
            df.to_excel(writer, sheet_name=category[:31], index=False)

print(f"\n✅ Done. File saved as: {output_file}")
