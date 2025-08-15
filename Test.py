import os
import time
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ================== CONFIG ==================
URL = "https://prod.marketrisk.citigroup.net/igm/#/FinancePL"   # adjust if your path differs
PAGE_LOAD_TIMEOUT = 120
TABLE_RENDER_TIMEOUT = 90
SCROLL_PAUSE = 0.8
MAX_SCROLLS = 200

# =============== SELENIUM HELPERS ===============
def start_driver():
    opts = Options()
    # Run with a visible browser so SSO can work:
    # opts.add_argument("--headless=new")  # <- enable only if your SSO allows headless
    opts.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver

def wait_visible(driver, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located(locator))

def wait_present(driver, locator, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))

def try_find(driver, by, sel, many=False):
    try:
        return (driver.find_elements if many else driver.find_element)(by, sel)
    except NoSuchElementException:
        return [] if many else None

def set_year_month(driver, yyyymm):
    """
    Try to populate the Year-Month control. Supports:
    - <input placeholder="Year-Month">
    - tokenized chips (it will clear existing chip if found)
    """
    # Clear existing chip/tag if present (close 'x')
    for chip_close_sel in ["button[aria-label='remove']", "span[aria-label='remove']", ".p-chip-remove-icon", ".ant-select-clear"]:
        for btn in try_find(driver, By.CSS_SELECTOR, chip_close_sel, many=True):
            try:
                btn.click(); time.sleep(0.2)
            except Exception:
                pass

    inp = (try_find(driver, By.CSS_SELECTOR, "input[placeholder*='Year'][placeholder*='Month']")
           or try_find(driver, By.XPATH, "//input[contains(@placeholder,'Year') and contains(@placeholder,'Month')]"))
    if not inp:
        # Some apps label it; try label association
        label = try_find(driver, By.XPATH, "//label[contains(.,'Year') and contains(.,'Month')]")
        if label:
            # next input sibling
            inp = try_find(driver, By.XPATH, "//label[contains(.,'Year') and contains(.,'Month')]/following::input[1]")
    if not inp:
        raise RuntimeError("Couldn't locate the Year-Month input field.")

    inp.click(); time.sleep(0.2)
    # If the control expects YYYYMM and auto-tokenizes on Enter:
    inp.clear()
    inp.send_keys(yyyymm)
    inp.send_keys("\n")
    time.sleep(0.5)

def click_search(driver):
    # Primary selector: exact button text
    candidates = [
        (By.XPATH, "//button[contains(., 'Search P_Value Statistics')]"),
        (By.XPATH, "//span[contains(., 'Search P_Value Statistics')]/ancestor::button[1]"),
        (By.CSS_SELECTOR, "button[title*='Search']"),
    ]
    for by, sel in candidates:
        btns = try_find(driver, by, sel, many=True)
        for b in btns:
            if b.is_displayed() and b.is_enabled():
                b.click()
                return
    raise RuntimeError("Couldn't find the 'Search P_Value Statistics' button.")

def set_page_size_all(driver):
    """
    Set page size to ALL. Works for <select> or custom dropdowns.
    """
    # 1) Native <select>
    selects = try_find(driver, By.CSS_SELECTOR, "select", many=True)
    for s in selects:
        try:
            Select(s).select_by_visible_text("ALL")
            time.sleep(0.8)
            return
        except Exception:
            pass

    # 2) Common custom dropdowns near the grid footer/header
    # Try opening the dropdown that currently shows 15/25/50…
    dropdown_candidates = [
        (By.XPATH, "//*[contains(@class,'page') or contains(@class,'paginator') or contains(@class,'size')][.//text()[contains(.,'15')]]"),
        (By.XPATH, "//button[contains(.,'15') or contains(.,'Rows') or contains(.,'per page')]"),
        (By.CSS_SELECTOR, "[role='combobox']"),
    ]
    opened = False
    for by, sel in dropdown_candidates:
        els = try_find(driver, by, sel, many=True)
        for e in els:
            try:
                if e.is_displayed(): 
                    e.click(); time.sleep(0.3); opened = True; break
            except Exception:
                continue
        if opened: 
            break

    if opened:
        # pick ALL from the popup/list
        all_opts = []
        all_opts += try_find(driver, By.XPATH, "//*[self::li or self::div or self::span][normalize-space(.)='ALL']", many=True)
        all_opts += try_find(driver, By.XPATH, "//*[contains(@class,'option') and normalize-space(.)='ALL']", many=True)
        for opt in all_opts:
            try:
                if opt.is_displayed():
                    opt.click(); time.sleep(1.0); return
            except Exception:
                pass

    # As a last resort: try to click any element with text ALL
    any_all = try_find(driver, By.XPATH, "//*[normalize-space(.)='ALL']", many=True)
    for a in any_all:
        try:
            if a.is_displayed():
                a.click(); time.sleep(1.0); return
        except Exception:
            pass

    # If we’re here, we couldn’t switch. Don’t fail—scraper will still paginate/scroll.
    print("Warning: couldn't set page size to ALL. Proceeding with pagination/scrolling.")

def wait_for_table(driver):
    """
    Wait until either a traditional <table> or an ARIA grid appears with rows.
    """
    wait = WebDriverWait(driver, TABLE_RENDER_TIMEOUT)
    wait.until(EC.any_of(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "[role='grid']"))
    ))
    # also wait for at least one data row
    wait.until(EC.any_of(
        EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "[role='row'] [role='gridcell']"))
    ))

def infinite_scroll_collect(driver, container):
    prev = 0
    for _ in range(MAX_SCROLLS):
        rows_now = len(container.find_elements(By.CSS_SELECTOR, "[role='row'], tbody tr, tr"))
        if rows_now == prev:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight;", container)
            time.sleep(SCROLL_PAUSE)
            rows_after = len(container.find_elements(By.CSS_SELECTOR, "[role='row'], tbody tr, tr"))
            if rows_after == prev:
                break
        else:
            prev = rows_now
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].clientHeight;", container)
            time.sleep(SCROLL_PAUSE)

def extract_from_html_table(table_el):
    headers = [h.text.strip() for h in table_el.find_elements(By.CSS_SELECTOR, "thead th")]
    if not headers:
        # sometimes headers are the first row
        first_tr = try_find(table_el, By.CSS_SELECTOR, "tr")
        if first_tr:
            headers = [th.text.strip() for th in first_tr.find_elements(By.CSS_SELECTOR, "th")]
    rows = []
    body_rows = table_el.find_elements(By.CSS_SELECTOR, "tbody tr")
    if not body_rows:
        body_rows = table_el.find_elements(By.CSS_SELECTOR, "tr")[1:]  # skip header
    for r in body_rows:
        tds = r.find_elements(By.CSS_SELECTOR, "td")
        if tds:
            rows.append([c.text.strip() for c in tds])
    return headers, rows

def extract_from_aria_grid(grid_el):
    headers = [h.text.strip() or f"col_{i+1}" for i, h in enumerate(grid_el.find_elements(By.CSS_SELECTOR, "[role='columnheader']"))]
    rows = []
    for r in grid_el.find_elements(By.CSS_SELECTOR, "[role='row']"):
        # skip header-rows
        if r.find_elements(By.CSS_SELECTOR, "[role='columnheader']"):
            continue
        cells = r.find_elements(By.CSS_SELECTOR, "[role='gridcell'],[role='cell']")
        if cells:
            rows.append([c.text.strip() for c in cells])
    return headers, rows

def scrape_table(driver):
    wait_for_table(driver)

    # try to identify the root (table or grid)
    table = try_find(driver, By.CSS_SELECTOR, "table")
    grid = try_find(driver, By.CSS_SELECTOR, "[role='grid']")
    root = table or grid
    if not root:
        raise RuntimeError("No table or grid found.")

    # attempt to load all rows via infinite scroll if a scrollable container exists
    scroll_container = root
    try:
        scroll_container = root.find_element(By.XPATH, ".//*[self::div or self::section][contains(@style,'overflow') or contains(@class,'scroll')][1]")
    except Exception:
        pass
    try:
        infinite_scroll_collect(driver, scroll_container)
    except Exception:
        pass

    if table:
        headers, rows = extract_from_html_table(table)
    else:
        headers, rows = extract_from_aria_grid(grid)

    if not rows:
        raise RuntimeError("No rows extracted. Try increasing timeouts or check selectors.")
    # normalize
    width = max(len(r) for r in rows)
    rows = [r + [""]*(width-len(r)) for r in rows]
    if not headers or len(headers) != width:
        headers = [f"col_{i+1}" for i in range(width)]
    df = pd.DataFrame(rows, columns=headers)

    # Nice column names if present
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
    # only rename those we actually have
    existing = {k: v for k, v in rename_map.items() if k in df.columns}
    df = df.rename(columns=existing)
    return df

def run_scrape(yyyymm: str, output_path: Path, log_fn=print):
    if not (len(yyyymm) == 6 and yyyymm.isdigit()):
        raise ValueError("Year-Month must be in YYYYMM format, e.g. 202504")

    driver = start_driver()
    try:
        driver.get(URL)
        # Wait for the app shell or table area
        try:
            WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Finance P&L')]")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table, [role='grid']"))
                )
            )
        except TimeoutException:
            log_fn("If SSO is required, complete the login in the opened browser window.")
            # give extra manual time
            messagebox.showinfo("Login", "Complete SSO/login in the opened browser, then click OK.")
        
        # Fill the year-month and search
        log_fn("Setting Year-Month…")
        set_year_month(driver, yyyymm)

        log_fn("Clicking Search…")
        click_search(driver)

        # Wait for results to render
        log_fn("Waiting for results…")
        wait_for_table(driver)

        # Set page size to ALL (if available)
        log_fn("Switching page size to ALL…")
        set_page_size_all(driver)

        # Scrape table
        log_fn("Scraping table…")
        df = scrape_table(driver)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        df.to_excel(output_path, index=False)
        log_fn(f"Saved {len(df):,} rows to:\n{output_path}")
        return True
    finally:
        try:
            driver.quit()
        except Exception:
            pass

# =============== TKINTER GUI ===============
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Finance P&L P-Value Scraper")
        self.geometry("520x220")
        self.resizable(False, False)

        container = ttk.Frame(self, padding=14)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Year-Month (YYYYMM):").grid(row=0, column=0, sticky="w")
        self.var_yyyymm = tk.StringVar()
        self.entry_yyyymm = ttk.Entry(container, textvariable=self.var_yyyymm, width=20)
        self.entry_yyyymm.grid(row=0, column=1, sticky="w", padx=(8,0))
        # default to current YYYYMM
        self.var_yyyymm.set(datetime.now().strftime("%Y%m"))

        self.run_btn = ttk.Button(container, text="Run and Save…", command=self.on_run)
        self.run_btn.grid(row=0, column=2, padx=(16,0))

        self.log = tk.Text(container, height=7, width=62, state="disabled")
        self.log.grid(row=1, column=0, columnspan=3, pady=(12,0), sticky="nsew")

        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(1, weight=1)

    def log_print(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def on_run(self):
        yyyymm = self.var_yyyymm.get().strip()
        if not (len(yyyymm) == 6 and yyyymm.isdigit()):
            messagebox.showerror("Invalid date", "Please enter Year-Month in YYYYMM format, e.g. 202504.")
            return

        # Default location: user's Documents
        documents = Path.home() / "Documents"
        documents.mkdir(exist_ok=True)
        default_name = f"{yyyymm}_pvalue.xlsx"
        initialfile = default_name
        initialdir = str(documents)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialdir=initialdir,
            initialfile=initialfile,
            title="Save extracted P-Values"
        )
        if not save_path:
            return
        out_path = Path(save_path)

        # Disable button during run
        self.run_btn.config(state="disabled")
        self.log_print(f"Starting… Output will be saved to:\n{out_path}")

        def worker():
            try:
                ok = run_scrape(yyyymm, out_path, log_fn=self.log_print)
                if ok:
                    self.log_print("Done.")
                    messagebox.showinfo("Success", f"Saved to:\n{out_path}")
            except Exception as e:
                self.log_print(f"Error: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    App().mainloop()
