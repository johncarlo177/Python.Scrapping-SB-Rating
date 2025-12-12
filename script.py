import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager

ua = UserAgent()
USER_AGENT = ua.random
ChromeDriverPath = "C:/chromedriver/chromedriver.exe"

BASE_URL = 'https://www.sportsbet.com.au'
FILE_NAME = 'Race Meetings.xlsm'
target_column = 23
ALLOWED_MEETINGS = ['(VIC)', '(NSW)', '(QLD)', '(SA)', '(WA)', '(NT)', '(TAS)', '(ACT)', '(NZ)', '(NZL)']
FS = {}
SR = {}

def setup_driver():
    options = Options()
    options.headless = True
    options.add_argument("--disable-images")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument(f"--user-agent={USER_AGENT}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-site-isolation-trials")

    service = Service(ChromeDriverPath)
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(800)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def extract_sb_rating(driver, race_url):
    global SR

    driver.get(BASE_URL + race_url)
    wait = WebDriverWait(driver, 15)

    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div[data-automation-id^='racecard-outcome-']")
        )
    )

    runner_ids = [
        el.get_attribute("data-automation-id").replace("racecard-outcome-", "")
        for el in driver.find_elements(
            By.CSS_SELECTOR,
            "div[data-automation-id^='racecard-outcome-']")
    ]

    print("--Next Race--")

    for runner_id in runner_ids:
        try:
            r = driver.find_element(
                By.CSS_SELECTOR,
                f"div[data-automation-id='racecard-outcome-{runner_id}']"
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", r
            )
            driver.execute_script("arguments[0].click();", r)

            soup = BeautifulSoup(driver.page_source, "html.parser")

            horse_el = soup.select_one(
                f"div[data-automation-id='racecard-outcome-{runner_id}'] "
                "div[data-automation-id='racecard-outcome-name'] span"
            )
            if not horse_el:
                continue

            raw_name = horse_el.get_text(strip=True)
            horse_name = re.sub(r"^\d+\.\s*", "", raw_name)

            sb_el = soup.select_one(
                f"div[data-automation-id='shortform-{runner_id}'] "
                "div[data-automation-id='shortform-SB Rating'] span:last-child"
            )
            if not sb_el:
                continue

            sb_rating = sb_el.get_text(strip=True)

            SR.setdefault("RACE", {})
            SR["RACE"][horse_name] = sb_rating

            print(f"✅ {horse_name} → SB Rating {sb_rating}")

        except Exception:
            # intentionally silent – DOM instability
            continue

def get_races(driver):
    driver.get(BASE_URL + "/racing-schedule")
    time.sleep(3)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # find all event cells
    cells = soup.find_all(
        "td",
        attrs={
            "data-automation-id": re.compile(
                r"^horse-racing-section-row-\d+-col-\d+-event-cell$"
            )
        },
    )

    race_links = []

    for td in cells:
        a = td.find("a", href=True)
        if a:
            race_links.append(a["href"])

    return race_links


def merge_excel(excel_file, FS):
    print("\n==============================")
    print("🔍 DEBUG: Starting merge_excel")
    print("==============================")

    workbook = load_workbook(filename=excel_file, keep_vba=True)

    def normalize(name):
        return name.strip().lower().replace("-", " ")

    # Normalize all sheet names
    normalized_sheet_map = {normalize(name): name for name in workbook.sheetnames}

    print("\n📄 Sheets in workbook:")
    for k, v in normalized_sheet_map.items():
        print(f"  '{k}'  →  '{v}'")

    print("\n📌 FS meetings loaded:", list(FS.keys()))
    print("📌 SR meetings loaded:", list(SR.keys()))

    # --- PROCESS SKY RATING ---
    print("\n==============================")
    print("🔸 PROCESSING SKY RATING (Col X)")
    print("==============================")

    for raw_sheet_name, horses in SR.items():
        norm_name = normalize(raw_sheet_name)
        actual_sheet_name = normalized_sheet_map.get(norm_name)

        print(f"\n➡ Meeting SR: '{raw_sheet_name}' normalized to '{norm_name}'")

        if not actual_sheet_name:
            print(f"❌ No matching sheet found for SR meeting: {raw_sheet_name}")
            continue

        print(f"✔ Matched SR sheet: {actual_sheet_name}")
        print(f"🌟 Horses in SR for this meeting: {list(horses.keys())}")

        sheet = workbook[actual_sheet_name]

        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                horse_name = str(cell.value).strip() if cell.value else ""
                if horse_name in horses:
                    sky_value = horses[horse_name]
                    sheet.cell(row=cell.row, column=24, value=sky_value)

                    print(f"   ⭐ Sky Saved | Row {cell.row} | Horse: '{horse_name}' | Value: {sky_value}")

    workbook.save(excel_file)
    print("\n==============================")
    print("🎉 Excel updated successfully")
    print("==============================\n")


def main():
    driver = setup_driver()

    race_links = get_races(driver)

    for race_link in race_links:
         extract_sb_rating(driver, race_link)

    # merge_excel(FILE_NAME, FS)  # FS optional
    # merge_excel(FILE_NAME, SR)  # save SB rating to column X

    driver.quit()

if __name__ == '__main__':
    main()
