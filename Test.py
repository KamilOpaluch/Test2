# pvalue_scraper_gui.py
import time
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ------------ CONFIG ------------
URL = "https://prod.marketrisk.citigroup.net/igm/#/FinancePL"  # adjust if your path differs
PAGE_LOAD_TIMEOUT = 120
TABLE_RENDER_TIMEOUT = 90
SCROLL_PAUSE = 0.8
MAX_SCROLLS = 200

# ------------ SELENIUM HELPERS ------------
def start_driver():
    opts = Options()
    # Keep visible for SSO; uncomment if headless is allowed in your environment:
    # opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver

def try_find(ctx, by, sel, many=False):
    try:
        return (ctx.find_elements if many else ctx.find_element)(by, sel)
    except NoSuchElementException:
        return [] if many else None

def wait_for_table(driver):
    wait = WebDriverWait(driver, TABLE_RENDER_TIMEOUT)
    wait.until(EC.any_of(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "[role='grid']"))
    ))
    wait.until(EC.any_of(
        EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "[role='row'] [role='gridcell']"))
    ))

def set_year_month(driver, yyyymm):
    # Clear existing chips/tokens if present
    for close_sel in ["button[aria-label='remove']", ".p-chip-remove-icon", ".ant-select-clear"]:
        for b in try_find(driver, By.CSS_SELECTOR, close_sel, many=True):
            try:
                b.click()
                time.sleep(0.2)
            except Exception:
                pass

    inp = (try_find(driver, By.CSS_SELECTOR, "input[placeholder*='Year'][placeholder*='Month']")
           or try_find(driver, By.XPATH, "//input[contains(@placeholder,'Year') and contains(@placeholder,'Month')]")
           or try_find(driver, By.XPATH, "//label[contains(.,'Year') and contains(.,'Month')]/following::input[1]"))
    if not inp:
        raise RuntimeError("Year-Month input not found.")
    inp.click()
    time.sleep(0.2)
    try:
        inp.clear()
    except Exception:
        pass
    inp.send_keys(yyyymm + "\n")
    time.sleep(0.5)

def click_search(driver):
    candidates = [
        (By.XPATH, "//button[contains(., 'Search P_Value Statistics')]"),
        (By.XPATH, "//span[contains(., 'Search P_Value Statistics')]/ancestor::button[1]"),
        (By.CSS_SELECTOR, "button[title*='Search']")
    ]
    for by, sel in candidates:
        btns = try_find(driver, by, sel, many=True)
        for b in btns:
            if b.is_displayed() and b.is_enabled():
                b.click()
                return
    raise RuntimeError("Search button not found.")

def set_page_size_all(driver):
    # 1) Native <select>
    for s in try_find(driver, By.CSS_SELECTOR, "select", many=True):
        try:
            Select(s).select_by_visible_text("ALL")
            time.sleep(0.8)
            return
        except Exception:
            pass

    # 2) Custom dropdowns
    opened = False
    for by, sel in [
        (By.XPATH, "//*[contains(@class,'page') or contains(@class,'paginator') or contains(@class,'size')][.//text()[contains(.,'15')]]"),
        (By.XPATH, "//button[contains(.,'15') or contains(.,'Rows') or contains(.,'per page')]"),
        (By.CSS_SELECTOR, "[role='combobox']")
    ]:
        els = try_find(driver, by, sel, many=True)
        for e in els:
            try:
                if e.is_displayed():
                    e.click()
                    time.sleep(0.3)
                    opened = True
                    break
            except Exception:
                pass
        if opened:
            break

    if opened:
        for opt in try_find(driver, By.XPATH, "//*[self::li or self::div or self::span][normalize-space(.)='ALL']", many=True):
            try:
                if opt.is_displayed():
                    opt.click()
                    time.sleep(1.0)
                    return
            except Exception:
                pass

    # 3) Last resort: click any visible 'ALL'
    for a in try_find(driver, By.XPATH, "//*[normalize-space(.)='ALL']", many=True):
        try:
            if a.is_displayed():
                a.click()
                time.sleep(1.0)
                return
        except Exception:
            pass

    print("Warning: couldn't switch page size to ALL; proceeding anyway.")

def extract_from_html_table(table_el):
    headers = [h.text.strip() for h in table_el.find_elements(By.CSS_SELECTOR, "thead th")]
    if not headers:
        first = try_find(table_el, By.CSS_SELECTOR, "tr")
        if first:
            headers = [th.text.strip() for th in first.find_elements(By.CSS_SELECTOR, "th")]
    rows = []
    for r in table_el.find_elements(By.CSS_SELECTOR, "tbody tr"):
        tds = r.find_elements(By.CSS_SELECTOR, "td")
        if tds:
            rows.append([c.text.strip() for c in tds])
    return headers, rows

def extract_from_aria_grid(grid_el):
    headers = [h.text.strip() or f"col_{i+1}" for i, h in enumerate(grid_el.find_elements(By.CSS_SELECTOR, "[role='columnheader']"))]
    rows = []
    for r in grid_el.find_elements(By.CSS_SELECTOR, "[role='row']"):
        if r.find_elements(By.CSS_SELECTOR, "[role='columnheader']"):
            continue
        cells = r.find_elements(By.CSS_SELECTOR, "[role='gridcell'],[role='cell']")
        if cells:
            rows.append([c.text.strip() for c in cells])
    return headers, rows

def scrape_table(driver):
    wait_for_table(driver)
    table = try_find(driver, By.CSS_SELECTOR, "table")
    grid = try_find(driver, By.CSS_SELECTOR, "[role='grid']")
    root = table or grid
    if not root:
        raise RuntimeError("Table/grid not found.")

    if table:
        headers, rows = extract_from_html_table(table)
    else:
        headers, rows = extract_from_aria_grid(grid)

    if not rows:
        raise RuntimeError("No rows extracted.")

    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    if not headers or len(headers) != width:
        headers = [f"col_{i+1}" for i in range(width)]

    df = pd.DataFrame(rows, columns=headers)
    # Friendly names if present
    rename = {
        "COB": "COB",
        "Level Code": "Level Code",
        "Level Id": "Level ID",
        "Region": "Region",
        "Ntr Request Id": "NTR Request ID",
        "P Value": "P Value",
        "User Stamp": "User",
        "Time Stamp": "Timestamp",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

def run_scrape(yyyymm: str, output_path: Path, log=print):
    if not (len(yyyymm) == 6 and yyyymm.isdigit()):
        raise ValueError("Year-Month must be YYYYMM, e.g. 202504.")

    driver = start_driver()
    try:
        driver.get(URL)
        # Wait for shell/table or pause for SSO
        try:
            WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
                EC.any_of(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Finance P&L')]")),
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table, [role='grid']"))
                )
            )
        except TimeoutException:
            messagebox.showinfo("Login required", "Complete SSO/login in the opened browser, then click OK.")

        log("Setting Year-Month…"); set_year_month(driver, yyyymm)
        log("Searching…"); click_search(driver)
        log("Waiting for results…"); wait_for_table(driver)
        log("Switching rows to ALL…"); set_page_size_all(driver)
        log("Scraping…")
        df = scrape_table(driver)
        df.to_excel(output_path, index=False)
        log(f"Saved {len(df):,} rows to:\n{output_path}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

# -------------------- TKINTER GUI --------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Finance P&L P-Value Scraper")
        self.geometry("520x220")
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Year-Month (YYYYMM):").grid(row=0, column=0, sticky="w")
        self.var_ym = tk.StringVar(value=datetime.now().strftime("%Y%m"))
        ttk.Entry(frame, textvariable=self.var_ym, width=20).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.run_btn = ttk.Button(frame, text="Run and Save…", command=self.on_run)
        self.run_btn.grid(row=0, column=2, padx=(16, 0))

        self.log = tk.Text(frame, height=7, width=62, state="disabled")
        self.log.grid(row=1, column=0, columnspan=3, pady=(12, 0))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(1, weight=1)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def on_run(self):
        yyyymm = self.var_ym.get().strip()
        if not (len(yyyymm) == 6 and yyyymm.isdigit()):
            messagebox.showerror("Invalid date", "Use YYYYMM, e.g. 202504.")
            return

        documents = Path.home() / "Documents"
        documents.mkdir(exist_ok=True)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialdir=str(documents),
            initialfile=f"{yyyymm}_pvalue.xlsx",
            title="Save extracted P-Values",
        )
        if not save_path:
            return

        out = Path(save_path)
        self.run_btn.config(state="disabled")
        self._log(f"Starting… Saving to:\n{out}")

        def worker():
            try:
                run_scrape(yyyymm, out, log=self._log)
                messagebox.showinfo("Success", f"Saved to:\n{out}")
            except Exception as e:
                self._log(f"Error: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    App().mainloop()
