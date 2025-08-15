import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# ==== CONFIG ====
URL = "https://prod.marketrisk.citigroup.net/igm/#/FinancePL"  # use your exact path
OUTPUT_DIR = Path.cwd()  # change if you want
PAGE_LOAD_TIMEOUT = 90
TABLE_RENDER_TIMEOUT = 60
SCROLL_PAUSE = 0.8   # for infinite-scroll style grids
MAX_SCROLLS = 200    # safety cap
# Optional: if the page needs a specific month, you can automate clicks later—start by scraping what's visible.

def start_driver():
    opts = Options()
    # Comment out next line if you want headless (SSO usually dislikes headless)
    # opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("detach", False)
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver

def wait_for_table(driver):
    """
    Wait for either a traditional <table> or an ARIA grid to appear.
    """
    wait = WebDriverWait(driver, TABLE_RENDER_TIMEOUT)
    try:
        locator = EC.any_of(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table")),
            EC.presence_of_element_located((By.CSS_SELECTOR, "[role='grid']"))
        )
        elem = wait.until(locator)
        return elem
    except TimeoutException:
        raise TimeoutException("Couldn't find a table or grid on the page. Check that you're on the right tab and logged in.")

def extract_from_html_table(table_el):
    # headers
    headers = []
    header_els = table_el.find_elements(By.CSS_SELECTOR, "thead th")
    if not header_els:
        header_els = table_el.find_elements(By.CSS_SELECTOR, "tr th")
    if header_els:
        headers = [h.text.strip() for h in header_els]

    rows = []
    body_rows = table_el.find_elements(By.CSS_SELECTOR, "tbody tr")
    if not body_rows:
        body_rows = table_el.find_elements(By.CSS_SELECTOR, "tr")[1:]  # skip header row fallback

    for r in body_rows:
        cells = r.find_elements(By.CSS_SELECTOR, "td")
        if not cells:
            continue
        row = [c.text.strip() for c in cells]
        rows.append(row)

    # align lengths
    if headers and rows and len(headers) != len(rows[0]):
        # fallback: make generic headers
        headers = [f"col_{i+1}" for i in range(len(rows[0]))]
    return headers, rows

def extract_from_aria_grid(grid_el):
    """
    For ag-Grid/PrimeNG: rows often have role='row' and cells role='gridcell'.
    """
    # headers
    headers = []
    header_cells = grid_el.find_elements(By.CSS_SELECTOR, "[role='columnheader']")
    if header_cells:
        headers = [h.text.strip() or f"col_{i+1}" for i, h in enumerate(header_cells)]

    # rows
    rows = []
    row_els = grid_el.find_elements(By.CSS_SELECTOR, "[role='row']")
    for r in row_els:
        # skip header rows that also use role='row'
        if r.find_elements(By.CSS_SELECTOR, "[role='columnheader']"):
            continue
        cells = r.find_elements(By.CSS_SELECTOR, "[role='gridcell'],[role='cell']")
        if cells:
            rows.append([c.text.strip() for c in cells])

    # align
    if headers and rows and len(headers) != len(rows[0]):
        headers = [f"col_{i+1}" for i in range(len(rows[0]))]
    return headers, rows

def try_click_next(driver):
    """
    Try to paginate using typical next controls. Returns True if it clicked and page should be re-scraped.
    """
    candidates = [
        ("button[aria-label*='Next']", By.CSS_SELECTOR),
        ("[aria-label='Next Page']", By.CSS_SELECTOR),
        ("button:has(svg[aria-label='Next'])", By.CSS_SELECTOR),
        ("//button[contains(., 'Next') or contains(., '›') or contains(., '>')]", By.XPATH),
        ("//a[contains(., 'Next') or contains(., '›') or contains(., '>')]", By.XPATH),
    ]
    for selector, how in candidates:
        try:
            if how == By.CSS_SELECTOR:
                btns = driver.find_elements(By.CSS_SELECTOR, selector)
            else:
                btns = driver.find_elements(By.XPATH, selector)
            for b in btns:
                label = (b.get_attribute("aria-disabled") or "").lower()
                disabled = label == "true" or "disabled" in (b.get_attribute("class") or "").lower() or not b.is_enabled()
                if not disabled and b.is_displayed():
                    b.click()
                    time.sleep(1.2)
                    return True
        except Exception:
            continue
    return False

def infinite_scroll_collect(driver, container):
    """
    If the grid uses infinite scroll/virtualization, scroll until no new rows appear.
    """
    prev_count = 0
    scrolls = 0
    while scrolls < MAX_SCROLLS:
        rows_now = len(container.find_elements(By.CSS_SELECTOR, "[role='row'], tbody tr, tr"))
        if rows_now == prev_count:
            # try to scroll a bit
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight;", container)
            time.sleep(SCROLL_PAUSE)
            rows_after = len(container.find_elements(By.CSS_SELECTOR, "[role='row'], tbody tr, tr"))
            if rows_after == prev_count:
                break
        else:
            prev_count = rows_now
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight;", container)
            time.sleep(SCROLL_PAUSE)
        scrolls += 1

def scrape_all_pages(driver):
    """
    Scrape all pages (handles table or aria grid; tries both; handles pagination and infinite scroll).
    """
    root = wait_for_table(driver)

    # if the grid is inside a scrollable container, keep a reference for infinite scroll
    scroll_container = root
    try:
        # prefer the nearest scrollable container if present
        scroll_container = root.find_element(By.XPATH, ".//*[self::div or self::section][contains(@style,'overflow') or contains(@class,'scroll')][1]")
    except Exception:
        pass

    # First, try to load everything via infinite scroll (harmless if not used)
    try:
        infinite_scroll_collect(driver, scroll_container)
    except Exception:
        pass

    all_headers = None
    all_rows = []

    def collect_once():
        nonlocal all_headers, all_rows
        # attempt HTML table
        table_els = driver.find_elements(By.CSS_SELECTOR, "table")
        if table_els:
            headers, rows = extract_from_html_table(table_els[0])
        else:
            # aria grid
            grid_els = driver.find_elements(By.CSS_SELECTOR, "[role='grid']")
            if not grid_els:
                raise RuntimeError("No table or grid found during collection.")
            headers, rows = extract_from_aria_grid(grid_els[0])

        if headers and not all_headers:
            all_headers = headers
        if rows:
            all_rows.extend(rows)

    # collect current page
    collect_once()

    # go through paginated pages if available
    seen_pages_safety = 0
    while try_click_next(driver):
        seen_pages_safety += 1
        if seen_pages_safety > 500:  # hard cap
            break
        # wait a bit for page to refresh
        time.sleep(1.0)
        # scroll if needed to force render
        try:
            infinite_scroll_collect(driver, scroll_container)
        except Exception:
            pass
        collect_once()

    # Build DataFrame, normalize sizes if headers are missing
    if not all_rows:
        raise RuntimeError("No rows collected. Check selectors or ensure the table is populated.")
    # normalize width
    max_len = max(len(r) for r in all_rows)
    norm_rows = [r + [""] * (max_len - len(r)) for r in all_rows]
    if not all_headers or len(all_headers) != max_len:
        all_headers = [f"col_{i+1}" for i in range(max_len)]
    df = pd.DataFrame(norm_rows, columns=all_headers)
    return df

def main():
    driver = start_driver()
    try:
        driver.get(URL)

        # ---- MANUAL SSO NOTE ----
        # If your company uses SSO, pause here until you're fully logged in and the page shows the table.
        try:
            # Wait until either the Finance P&L page label or table appears.
            WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Finance P&L')]")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table, [role='grid']"))
                )
            )
        except TimeoutException:
            print("Timed out waiting for the app to load. If SSO is required, complete it in the opened browser.")
            # Give user extra time to complete SSO, then continue
            input("Press Enter after you see the Finance P&L table...")

        df = scrape_all_pages(driver)

        # Optional tidy-up: remove empty columns, strip whitespace
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        # Example: ensure expected columns exist; if they do, rename for neatness
        rename_map = {
            "COB": "COB",
            "Level Code": "Level Code",
            "Level Id": "Level ID",
            "Region": "Region",
            "Ntr Request Id": "NTR Request ID",
            "P Value": "P Value",
            "User Stamp": "User",
            "Time Stamp": "Timestamp",
        }
        for k, v in list(rename_map.items()):
            if k not in df.columns:
                rename_map.pop(k, None)
        df = df.rename(columns=rename_map)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = OUTPUT_DIR / f"finance_pnl_{ts}.xlsx"
        df.to_excel(out_path, index=False)
        print(f"Saved {len(df):,} rows to {out_path}")

    finally:
        # Close the browser
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
