from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller
import time
import pandas as pd
import re

BASE_URL = "https://www.booksellers.org.uk/bookshopsearch"
DETAIL_PATTERN = "Bookshop-Details.aspx?m="
LOAD_MORE_XPATH = (
    '//input[contains(translate(@value, "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "LOAD MORE RESULTS")]'
    ' | //a[contains(translate(normalize-space(.), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "LOAD MORE RESULTS")]'
)
COOKIE_ACCEPT = '//a[contains(@class, "cc_btn_accept")]'


def create_driver():
    chromedriver_autoinstaller.install()
    opts = Options()
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=opts)


def js_click(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
        element,
    )


def extract_entries(driver):
    anchors = driver.find_elements(By.XPATH, f'//a[contains(@href, "{DETAIL_PATTERN}")]')
    seen = {}
    results = []
    for a in anchors:
        href = a.get_attribute("href")
        if href in seen:
            continue
        seen[href] = True
        results.append(href)
    return results


def scrape_all():
    driver = create_driver()
    seen_urls = set()
    all_links = []
    last_seen_count = 0

    try:
        driver.get(BASE_URL)
        time.sleep(2)

        try:
            cookie_btn = driver.find_element(By.XPATH, COOKIE_ACCEPT)
            js_click(driver, cookie_btn)
            time.sleep(2)
        except NoSuchElementException:
            pass

        search_btn = driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="Search"]')
        js_click(driver, search_btn)
        time.sleep(5)

        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(5)

            batch = extract_entries(driver)
            new_links = [u for u in batch if u not in seen_urls]
            all_links.extend(new_links)
            for u in new_links:
                seen_urls.add(u)
            print(f"🔎 Collected so far: {len(seen_urls)}")

            last_seen_count = len(seen_urls)

            try:
                btn = driver.find_element(By.XPATH, LOAD_MORE_XPATH)
                js_click(driver, btn)
                time.sleep(15)  # Updated wait time to 15 seconds
            except NoSuchElementException:
                print("✅ No more Load More button. Scraping complete.")
                break

    finally:
        driver.quit()

    with open("urls.txt", "w", encoding="utf-8") as f:
        for url in all_links:
            f.write(url + "\n")
    print(f"\n📝 Saved {len(all_links)} URLs → urls.txt")


if __name__ == "__main__":
    scrape_all()
