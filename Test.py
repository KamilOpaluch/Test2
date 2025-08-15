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
            if a.is_displayed():_
