import pandas as pd
import re
import sys
import logging
from pathlib import Path

# -----------------------
# Logging setup
# -----------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger("pnl-pipeline")

# -----------------------
# Config / Paths
# -----------------------
base = Path.home() / "Documents" / "IMA_Extend"
monthly_pattern = re.compile(r".*_(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.xlsx$", re.IGNORECASE)

USECOLS = ["LV ID","MSH Level 4","Firm Account","Date","Tot Clean P&L ex Theta","Actual","Rwa Type"]

# Canonical column names we will use everywhere
WANTED = [
    "MSH Level 4",
    "Firm Account",
    "Date",
    "Tot Clean P&L ex Theta",
    "Actual",
    "Rwa Type",
]

# -----------------------
# Helpers
# -----------------------
def normalize_cols(cols):
    return [re.sub(r"\s+", " ", str(c)).strip() for c in cols]

def normalize_account_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.strip()
         .str.replace(r"\.0$", "", regex=True)
    )

def ensure_numeric(df: pd.DataFrame, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def load_monthly_file(path: Path) -> pd.DataFrame:
    try:
        log.info("Reading monthly file: %s", path.name)
        df = pd.read_excel(
            path,
            header=11,  # row 12 are headers
            engine="openpyxl",
            usecols=USECOLS,
            dtype={"LV ID": str}
        )
        df.columns = normalize_cols(df.columns)
        log.info("[%s] Loaded %d rows, %d columns", path.name, df.shape[0], df.shape[1])

        # LV filter
        if "LV ID" in df.columns:
            lv = (
                df["LV ID"].astype(str).str.strip()
                  .str.replace(r"\.0$", "", regex=True)
                  .str.zfill(5)
            )
            before = len(df)
            df = df[lv == "06500"]
            log.info("[%s] LV ID filter kept %d / %d rows", path.name, len(df), before)
        else:
            log.warning("[%s] Column 'LV ID' not found; skipping LV filter", path.name)

        # Ensure all WANTED exist
        for c in WANTED:
            if c not in df.columns:
                df[c] = pd.NA

        out = df[WANTED].copy()
        out["Firm Account"] = normalize_account_series(out["Firm Account"])
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
        out = ensure_numeric(out, ["Tot Clean P&L ex Theta", "Actual"])
        out["Rwa Type"] = out["Rwa Type"].astype(str).str.strip()

        # Drop rows without Date or Account (most useful for daily time series)
        before = len(out)
        out = out.dropna(subset=["Date", "Firm Account"])
        log.info("[%s] Normalized & selected -> %d rows (dropped %d empty-date/account rows)",
                 path.name, len(out), before - len(out))
        return out
    except Exception as e:
        log.exception("Failed reading %s: %s", path.name, e)
        # Return empty with correct columns so pipeline continues
        return pd.DataFrame(columns=WANTED)

def load_monthlies(folder: Path) -> pd.DataFrame:
    files = [p for p in folder.glob("*.xlsx") if monthly_pattern.match(p.name)]
    log.info("Discovered %d monthly files in %s", len(files), folder)
    for p in files:
        log.info("  - %s", p.name)
    if not files:
        raise FileNotFoundError(f"No monthly *_Mon.xlsx files found in {folder}")

    frames = []
    for p in files:
        frames.append(load_monthly_file(p))
    if not frames:
        return pd.DataFrame(columns=WANTED)
    combined = pd.concat(frames, ignore_index=True)
    log.info("Monthly combined: %d rows, %d unique accounts", len(combined), combined["Firm Account"].nunique())
    return combined

def load_mtd_adjustments(folder: Path) -> pd.DataFrame:
    path = folder / "MTD_adjusmtents.xlsx"
    log.info("Reading MTD adjustments: %s (sheet='Adjustment')", path.name)
    try:
        df = pd.read_excel(path, sheet_name="Adjustment", engine="openpyxl")
    except Exception as e:
        log.exception("Failed to read %s: %s", path, e)
        return pd.DataFrame(columns=WANTED)

    df.columns = normalize_cols(df.columns)
    log.info("[MTD] Raw rows: %d, columns: %d", df.shape[0], df.shape[1])

    # Optional LV filter if present
    if "LV ID" in df.columns:
        lv = (
            df["LV ID"].astype(str).str.strip()
              .str.replace(r"\.0$", "", regex=True)
              .str.zfill(5)
        )
        before = len(df)
        df = df[lv == "06500"]
        log.info("[MTD] LV ID filter kept %d / %d rows", len(df), before)

    for c in WANTED:
        if c not in df.columns:
            df[c] = pd.NA

    out = df[WANTED].copy()
    out["Firm Account"] = normalize_account_series(out["Firm Account"])
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = ensure_numeric(out, ["Tot Clean P&L ex Theta", "Actual"])
    out["Rwa Type"] = out["Rwa Type"].astype(str).str.strip()

    kept = len(out.dropna(subset=["Firm Account"]))
    log.info("[MTD] Prepared rows: %d (Firm Account present), unique accounts: %d",
             kept, out["Firm Account"].nunique())
    return out.dropna(subset=["Firm Account"])

def load_mcc_accounts(folder: Path) -> pd.DataFrame:
    path = folder / "mcc.xlsx"
    log.info("Reading account map: %s ('new' & 'old')", path.name)

    def load_sheet(name):
        try:
            df = pd.read_excel(path, sheet_name=name, engine="openpyxl")
        except Exception as e:
            log.exception("Failed reading %s sheet '%s': %s", path.name, name, e)
            return pd.DataFrame(columns=["Firm Account", "LEGAL_ENTITY"])
        df.columns = normalize_cols(df.columns)
        need = ["LEVEL_12", "LEGAL_ENTITY"]
        for c in need:
            if c not in df.columns:
                df[c] = pd.NA
        out = df[need].copy()
        out["Firm Account"] = (
            out["LEVEL_12"].astype(str).str.strip()
            .str.replace(r"^S2", "", regex=True)
            .str.replace(r"\.0$", "", regex=True)
        )
        out["LEGAL_ENTITY"] = out["LEGAL_ENTITY"].astype(str).str.strip().str.upper()
        out = out[["Firm Account", "LEGAL_ENTITY"]].dropna(subset=["Firm Account"])
        log.info("[mcc:%s] rows=%d, unique accounts=%d", name, len(out), out["Firm Account"].nunique())
        return out

    new_df = load_sheet("new")
    old_df = load_sheet("old")

    combo = pd.concat([new_df, old_df], ignore_index=True)
    before_dupes = combo.shape[0]
    combo = combo.drop_duplicates(subset=["Firm Account"], keep="first")
    dupes_removed = before_dupes - combo.shape[0]
    combo = combo[combo["LEGAL_ENTITY"].isin(["CGML", "CGME"])].copy()
    log.info("[mcc] Combined unique accounts: %d (removed %d duplicates). CGML=%d, CGME=%d",
             combo.shape[0],
             dupes_removed,
             (combo["LEGAL_ENTITY"]=="CGML").sum(),
             (combo["LEGAL_ENTITY"]=="CGME").sum())
    return combo.reset_index(drop=True)

def timeseries(df: pd.DataFrame, entity: str, tb_only: bool) -> pd.DataFrame:
    sub = df[df["LEGAL_ENTITY"] == entity].copy()
    if tb_only:
        sub = sub[sub["Rwa Type"].astype(str).str.strip().str.upper() == "TB"]
    if sub.empty:
        log.warning("[TS] %s | TB=%s -> no rows", entity, tb_only)
        return pd.DataFrame(columns=["Date","Tot Clean P&L ex Theta","Actual"])
    agg = (
        sub.groupby("Date", as_index=False)[["Tot Clean P&L ex Theta", "Actual"]]
           .sum()
           .sort_values("Date", kind="stable")
    )
    dmin = sub["Date"].min()
    dmax = sub["Date"].max()
    log.info("[TS] %s | TB=%s -> %d dates (range: %s .. %s)",
             entity, tb_only, agg.shape[0], dmin.date() if pd.notna(dmin) else "NaT",
             dmax.date() if pd.notna(dmax) else "NaT")
    return agg

# -----------------------
# Orchestration
# -----------------------
log.info("=== Pipeline start ===")
log.info("Base folder: %s", base)

# 1) Load components
daily_raw = load_monthlies(base)
log.info("Daily raw: %d rows, %d unique accounts", len(daily_raw), daily_raw["Firm Account"].nunique())

mtd_raw = load_mtd_adjustments(base)
log.info("MTD raw: %d rows, %d unique accounts", len(mtd_raw), mtd_raw["Firm Account"].nunique())

mcc_map = load_mcc_accounts(base)

# 2) Filter to MCC accounts + join LEGAL_ENTITY
daily_df = daily_raw.merge(mcc_map, on="Firm Account", how="inner")
mtd_df   = mtd_raw.merge(mcc_map,   on="Firm Account", how="inner")

log.info("Daily after MCC filter: %d rows, %d accounts", len(daily_df), daily_df["Firm Account"].nunique())
log.info("MTD after MCC filter:   %d rows, %d accounts", len(mtd_df),   mtd_df["Firm Account"].nunique())

# Coverage diagnostics
missing_accounts = set(daily_raw["Firm Account"].dropna().unique()) - set(mcc_map["Firm Account"].dropna().unique())
if missing_accounts:
    log.warning("Accounts in daily not present in MCC map: %d (showing up to 5): %s",
                len(missing_accounts), sorted(list(missing_accounts))[:5])

# 3) Combine daily + adjustments (as additional rows)
all_df = pd.concat([daily_df, mtd_df], ignore_index=True)
all_df = all_df[WANTED + ["LEGAL_ENTITY"]]
all_df = all_df.sort_values(["LEGAL_ENTITY", "Date", "Firm Account"], kind="stable").reset_index(drop=True)
log.info("Combined total rows: %d | Unique dates: %d | Unique accounts: %d",
         len(all_df),
         all_df["Date"].nunique(),
         all_df["Firm Account"].nunique())

# 4) Build time series (All vs TB-only) for CGML & CGME
cgml_all = timeseries(all_df, "CGML", tb_only=False)
cgml_tb  = timeseries(all_df, "CGML", tb_only=True)
cgme_all = timeseries(all_df, "CGME", tb_only=False)
cgme_tb  = timeseries(all_df, "CGME", tb_only=True)

# 5) Write Excel output
out_path = base / "pnl_timeseries_summary.xlsx"
log.info("Writing Excel to: %s", out_path)
with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
    all_df.to_excel(xw, sheet_name="Combined_Rows", index=False)
    mcc_map.to_excel(xw, sheet_name="Account_Map", index=False)
    cgml_all.to_excel(xw, sheet_name="CGML_All", index=False)
    cgml_tb.to_excel(xw,  sheet_name="CGML_TB", index=False)
    cgme_all.to_excel(xw, sheet_name="CGME_All", index=False)
    cgme_tb.to_excel(xw,  sheet_name="CGME_TB", index=False)
log.info("Excel written successfully.")

log.info("=== Pipeline done ===")
