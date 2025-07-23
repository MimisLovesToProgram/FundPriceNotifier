#!/usr/bin/python3

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import time
import re
import datetime

today = datetime.datetime.today()
print(today)

links = ["https://www.blackrock.com/americas-offshore/en/products/229298/blackrock-new-energy-a2-eur-fund",
        "https://am.jpmorgan.com/gb/en/asset-management/adv/products/jpm-europe-strategic-value-a-dist-eur-lu0107398884#",
	"https://am.jpmorgan.com/gb/en/asset-management/adv/products/jpm-nasdaq-equity-premium-income-active-ucits-etf-usd-dist-ie000u9j8hx9",
        "https://markets.ft.com/data/funds/tearsheet/summary?s=LU1775978304:EUR",
        "https://www.franklintempleton.lu/our-funds/price-and-performance/products/2399/Z/templeton-european-opportunities-fund/LU0122612848",
        "https://www.franklintempleton.lu/our-funds/price-and-performance/products/4345/T/franklin-india-fund/LU0231205187",
        "https://www.franklintempleton.lu/our-funds/price-and-performance/products/2182/T/templeton-global-equity-income-fund/LU0211332647",
        "https://www.franklintempleton.lu/our-funds/price-and-performance/products/4343/T/templeton-bric-fund/LU0229946628",
	"https://www.tradingview.com/symbols/XETR-LYMS/"]

prices = []
excel_path = "/4TB_WD/Mimis/Prices/prices.xlsx"
price_pattern = r"[0-9]+\.[0-9]+$" # This matches what a price should look like, so that we never get something inconsistent if the site changes or something

# Ignoring superficial logs, running the browser with no graphics, and other minor details you definitely know (!).
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument('--log-level=3')
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--disable-infobars")
options.add_argument("--disable-notifications")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

service = Service("/usr/bin/chromedriver")

driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

for link in links:
    driver.get(link)
    time.sleep(10)

    if "blackrock" in link:
        # Accept necessary cookies
        locator = (By.ID, "onetrust-reject-all-handler")
        necessary = wait.until(EC.presence_of_element_located(locator))
        necessary.click()

        # Accept terms & conditions
        locator = (By.ID, "terms-and-conditions-masthead-accept")
        accept = wait.until(EC.presence_of_element_located(locator))
        accept.click()

        time.sleep(2)

        # First search for the fund because the whole link redirects to the product list, not the specific one (site's problem)
        locator = (By.ID, "screener-keyword-search-input")
        inputfield = wait.until(EC.presence_of_element_located(locator))
        inputfield.send_keys("LU0171289902") # The fund's code

        time.sleep(2)

        # Now get all divs with this class, from the one search result. These have the fund's info
        locator = (By.CSS_SELECTOR, "div.ng-star-inserted")
        elems = wait.until(EC.presence_of_all_elements_located(locator))

        nav = elems[8].text[4:] # NAV (price) is in the 8th div. We also want to exclude the "NAV\n" in the beggining

        if re.search(price_pattern, nav):
            prices.append(nav)
        else:
            prices.append(-1)

    elif "jpmorgan" in link:
        # Only after those buttons are pressed does the site load the NAV.
        # Also, you are only asked to press it once (doesn't appear in the second visit). ie000u9j8hx9 is the code of the second fund from jpmorgan.
        if not "ie000u9j8hx9" in link:
            locator = (By.CSS_SELECTOR, "button.cta-btn.secondary")
            uk_button = wait.until(EC.presence_of_element_located(locator))
            uk_button.click()

            locator = (By.ID, "onetrust-reject-all-handler")
            reject_btn = wait.until(EC.presence_of_element_located(locator))
            reject_btn.click()

            locator = (By.CSS_SELECTOR, "a.accept")
            accept_terms = wait.until(EC.presence_of_element_located(locator))
            accept_terms.click()

        # Then just get the NAV by the element with the argument data-testid set to [this]
        locator = (By.CSS_SELECTOR, '[data-testid="quick-stat-nav-value"]')
        nav_elem = wait.until(EC.presence_of_element_located(locator))

        nav = nav_elem.text[4:] # Also removing the "EUR " from the text, to be left with a numerical price

        if re.search(price_pattern, nav):
            prices.append(nav)
        else:
            prices.append(-1)

    elif "markets.ft" in link:
        # Like in BlackRock, there are many elements with the same tag and class as the one with the NAV price. Working similarly.
        locator = (By.CSS_SELECTOR, "span.mod-ui-data-list__value")
        elems = wait.until(EC.presence_of_all_elements_located(locator))

        nav = elems[0].text # NAV is the 1st span with this class

        if re.search(price_pattern, nav):
            prices.append(nav)
        else:
            prices.append(-1)

    elif "franklintempleton" in link:
        # No questions asked, just loading the site & getting the price
        locator = (By.CSS_SELECTOR, "p.summary-item__value.summary-item__value--basic")
        nav_elem = wait.until(EC.presence_of_element_located(locator))

        nav = nav_elem.text[1:] # Excluding "€" from the NAV text

        if re.search(price_pattern, nav):
            prices.append(nav)
        else:
            prices.append(-1)

    elif "tradingview" in link:
        # Same, just get the price
        locator = (By.CSS_SELECTOR, ".last-zoF9r75I.js-symbol-last")
        price_elem = wait.until(EC.presence_of_element_located(locator))

        price = price_elem.text
        if re.search(price_pattern, price):
            prices.append(price)
        else:
            prices.append(-1)

excel = load_workbook(excel_path)
sheet = excel.active

normal_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # White background cell fill
missing_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # Yellow background cell fill

for i, price in enumerate(prices, start=2):
    if price != -1:
        sheet[f"A{i}"].fill = normal_fill
        sheet[f"A{i}"] = float(price)
        sheet[f"D{i}"] = '=VALUE("' + today.strftime("%d/%m/%Y") + '")' # Too many quotes to use an f-string...
    else:
        sheet[f"A{i}"].fill = missing_fill # We don't change the text value, just inform that it is a past price & something went wrong here or on the site.

excel.save(excel_path)
print("")
