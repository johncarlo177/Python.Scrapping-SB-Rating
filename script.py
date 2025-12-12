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

BASE_URL = 'https://www.sportsbet.com.au/'
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

def extract_sb_rating(driver, url):
    global SR

    driver.get(BASE_URL + url)
    time.sleep(3)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # --- SELECT ALL RUNNERS ---
    runners = soup.find_all(
        "div",
        attrs={
            "data-automation-id": re.compile(r"^racecard-outcome-\d+$")
        }
    )

    for r in runners:
        try:
            # Horse name
            horse_el = r.select_one("div[data-automation-id='racecard-outcome-name'] span")
            if not horse_el:
                continue

            horse_name = horse_el.get_text(strip=True)
            
            # Horse ID used in shortform
            hid = r.get("data-automation-id").replace("racecard-outcome-", "")

            # Find shortform container by ID
            sf = soup.select_one(f"div[data-automation-id='shortform-{hid}']")
            if not sf:
                continue

            # Extract SB Rating
            sb_div = sf.select_one("div[data-automation-id='shortform-SB Rating']")
            if not sb_div:
                continue

            spans = sb_div.select("span")
            sb_rating = spans[-1].get_text(strip=True)   # last span contains the number

            # Save
            SR.setdefault("RACE", {})    # You can change name
            SR["RACE"][horse_name] = sb_rating

            print(f"RaceName, {horse_name}, SB Rating {sb_rating}")

        except Exception as e:
            print("Error:", e)
            continue


def get_race_links(driver, meeting_url):
    driver.get(BASE_URL + meeting_url)
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    links = []

    race_items = soup.select("a.link_fqiekv4")  # SportsBet uses this class for race links

    for a in race_items:
        href = a.get("href")
        if href and "/race-" in href:
            links.append(href)

    return links


def get_meetings(driver):
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

    meeting_links = []

    for td in cells:
        a = td.find("a", href=True)
        if a:
            meeting_links.append(a["href"])

    return meeting_links



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

    meeting_links = get_meetings(driver)

    for meeting_url in meeting_links:
        print("📌 Meeting:", meeting_url)

        race_links = extract_sb_rating(driver, meeting_url)

        for r_url in race_links:
            print(" → Race:", r_url)
            extract_sb_rating(driver, r_url, [])  # empty list = disable filtering

    merge_excel(FILE_NAME, FS)  # FS optional
    merge_excel(FILE_NAME, SR)  # save SB rating to column X

    driver.quit()

if __name__ == '__main__':
    main()
