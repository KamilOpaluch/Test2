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
    # Keep original face value (no zfill), trim, remove Excel ".0"
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

def make_acc_key(s: pd.Series) -> pd.Series:
    # Key used for joining: drop leading zeros so "00123" == "123"
    key = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    key = key.str.lstrip("0")
    # if everything was zeros, keep "0"
    key = key.mask(key.eq(""), "0")
    return key

def ensure_numeric(df: pd.DataFrame, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def entity_hint_from_msh4(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.upper()
    hint = pd.Series(pd.NA, index=t.index, dtype="object")
    hint = hint.mask(t.str.contains(r"\bCGML\b"), "CGML")
    hint = hint.mask(t.str.contains(r"\bCGME\b"), "CGME")
    return hint

# -----------------------
# Loaders
# -----------------------
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

        # Add join key + entity hint
        out["acc_key"] = make_acc_key(out["Firm Account"])
        out["entity_hint"] = entity_hint_from_msh4(out["MSH Level 4"])

        # Drop rows without Date or Account
        before = len(out)
        out = out.dropna(subset=["Date", "Firm Account"])
        log.info("[%s] Normalized & selected -> %d rows (dropped %d empty-date/account rows)",
                 path.name, len(out), before - len(out))
        return out
    except Exception as e:
        log.exception("Failed reading %s: %s", path.name, e)
        return pd.DataFrame(columns=WANTED + ["acc_key","entity_hint"])

def load_monthlies(folder: Path) -> pd.DataFrame:
    files = [p for p in folder.glob("*.xlsx") if monthly_pattern.match(p.name)]
    log.info("Discovered %d monthly files in %s", len(files), folder)
    for p in files:
        log.info("  - %s", p.name)
    if not files:
        raise FileNotFoundError(f"No monthly *_Mon.xlsx files found in {folder}")

    frames = [load_monthly_file(p) for p in files]
    if not frames:
        return pd.DataFrame(columns=WANTED + ["acc_key","entity_hint"])
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
        return pd.DataFrame(columns=WANTED + ["acc_key","entity_hint"])

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

    # Add join key + entity hint (even if not used much for MTD)
    out["acc_key"] = make_acc_key(out["Firm Account"])
    out["entity_hint"] = entity_hint_from_msh4(out["MSH Level 4"])

    kept = len(out.dropna(subset=["Firm Account"]))
    log.info("[MTD] Prepared rows: %d (Firm Account present), unique accounts: %d",
             kept, out["Firm Account"].nunique())
    return out.dropna(subset=["Firm Account"])

def load_mcc_accounts(folder: Path) -> pd.DataFrame:
    path = folder / "mcc.xlsx"
    log.info("Reading account map: %s ('new' & 'old')", path.name)

    def load_sheet(name, source_label):
        try:
            df = pd.read_excel(path, sheet_name=name, engine="openpyxl")
        except Exception as e:
            log.exception("Failed reading %s sheet '%s': %s", path.name, name, e)
            return pd.DataFrame(columns=["Firm Account", "LEGAL_ENTITY", "source", "acc_key"])
        df.columns = normalize_cols(df.columns)
        need = ["LEVEL_12", "LEGAL_ENTITY"]
        for c in need:
            if c not in df.columns:
                df[c] = pd.NA
        out = df[need].copy()
        # Remove S2 prefix (with optional separators/spaces), keep digits/letters after
        out["Firm Account"] = (
            out["LEVEL_12"].astype(str).str.strip()
            .str.replace(r"^S2[\s\-_]*", "", regex=True)
            .str.replace(r"\.0$", "", regex=True)
        )
        out["LEGAL_ENTITY"] = out["LEGAL_ENTITY"].astype(str).str.strip().str.upper()
        out["source"] = source_label
        out["acc_key"] = make_acc_key(out["Firm Account"])
        out = out[["Firm Account","acc_key","LEGAL_ENTITY","source"]].dropna(subset=["Firm Account"])
        # keep both CGML & CGME if present — no dedup here!
        log.info("[mcc:%s] rows=%d, unique accounts=%d", name, len(out), out["Firm Account"].nunique())
        return out

    new_df = load_sheet("new", "new")
    old_df = load_sheet("old", "old")
    combo = pd.concat([new_df, old_df], ignore_index=True)

    # Only CGML/CGME rows; keep duplicates (account can belong to both)
    combo = combo[combo["LEGAL_ENTITY"].isin(["CGML", "CGME"])].copy()
    log.info(
        "[mcc] Combined rows=%d | unique accounts=%d | CGML rows=%d | CGME rows=%d (duplicates allowed)",
        len(combo),
        combo["Firm Account"].nunique(),
        (combo["LEGAL_ENTITY"]=="CGML").sum(),
        (combo["LEGAL_ENTITY"]=="CGME").sum()
    )
    return combo.reset_index(drop=True)

# -----------------------
# Entity assignment logic (per-row resolution)
# -----------------------
def attach_entity_with_resolution(daily_like: pd.DataFrame, mcc_map: pd.DataFrame) -> pd.DataFrame:
    """
    For each row in daily_like, attach a single LEGAL_ENTITY using:
    1) entity_hint (from MSH Level 4) if matches a mapping row
    2) otherwise prefer mapping from 'new'
    3) otherwise take first available mapping row

    Returns daily_like with a 'LEGAL_ENTITY' column.
    """
    if daily_like.empty:
        daily_like["LEGAL_ENTITY"] = pd.NA
        return daily_like

    work = daily_like.reset_index(drop=False).rename(columns={"index":"_row_id"})
    merged = work.merge(mcc_map, on="acc_key", how="left", suffixes=("","_map"))

    # If no mapping at all, we will drop later but log first
    no_map = merged["LEGAL_ENTITY"].isna().groupby(merged["_row_id"]).all()
    missing_rows = no_map[no_map].index.tolist()
    if missing_rows:
        log.warning("Entity mapping missing for %d rows (sample row_ids: %s)",
                    len(missing_rows), missing_rows[:5])

    # Rank candidates per row_id
    merged["hint_match"] = merged["LEGAL_ENTITY"].eq(merged["entity_hint"])
    merged["source_rank"] = merged["source"].map({"new":0, "old":1}).fillna(2).astype(int)

    # Sort: hint matches first, then prefer 'new', then anything
    merged = merged.sort_values(["_row_id", "hint_match", "source_rank"], ascending=[True, False, True])

    # Keep best candidate per original row
    picked = merged.drop_duplicates(subset=["_row_id"], keep="first")

    # Diagnostics
    total = len(work)
    matched_hint = picked["hint_match"].sum()
    used_new = ((picked["hint_match"]==False) & (picked["source_rank"]==0)).sum()
    fallback = ((picked["hint_match"]==False) & (picked["source_rank"]>0)).sum()
    log.info("Entity resolution: %d rows | hint-matched=%d | prefer-new=%d | fallback=%d",
             total, matched_hint, used_new, fallback)

    out = picked.drop(columns=["_row_id","hint_match","source_rank","source"])
    return out

# -----------------------
# Time series builder
# -----------------------
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

# 2) Attach LEGAL_ENTITY per row via conflict-aware resolution
daily_df = attach_entity_with_resolution(daily_raw, mcc_map)
mtd_df   = attach_entity_with_resolution(mtd_raw,   mcc_map)

# Drop rows where entity could not be mapped
before_daily = len(daily_df)
before_mtd   = len(mtd_df)
daily_df = daily_df[daily_df["LEGAL_ENTITY"].isin(["CGML","CGME"])]
mtd_df   = mtd_df[mtd_df["LEGAL_ENTITY"].isin(["CGML","CGME"])]

log.info("Daily after entity attach: %d rows (dropped %d unmapped), accounts=%d, CGML=%d, CGME=%d",
         len(daily_df), before_daily - len(daily_df),
         daily_df["Firm Account"].nunique(),
         (daily_df["LEGAL_ENTITY"]=="CGML").sum(),
         (daily_df["LEGAL_ENTITY"]=="CGME").sum())

log.info("MTD after entity attach:   %d rows (dropped %d unmapped), accounts=%d, CGML=%d, CGME=%d",
         len(mtd_df), before_mtd - len(mtd_df),
         mtd_df["Firm Account"].nunique(),
         (mtd_df["LEGAL_ENTITY"]=="CGML").sum(),
         (mtd_df["LEGAL_ENTITY"]=="CGME").sum())

# 3) Combine daily + adjustments (as additional rows)
all_df = pd.concat([daily_df[WANTED + ["LEGAL_ENTITY"]],
                    mtd_df[WANTED + ["LEGAL_ENTITY"]]], ignore_index=True)

# Sort and tidy
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
    # Raw combined rows (with entity)
    all_df.to_excel(xw, sheet_name="Combined_Rows", index=False)

    # For traceability, dump unresolved or ambiguous diagnostics if needed
    # (Here we just include the union of MCC mappings)
    mcc_map.to_excel(xw, sheet_name="Account_Map", index=False)

    # Time series per entity & scenario
    cgml_all.to_excel(xw, sheet_name="CGML_All", index=False)
    cgml_tb.to_excel(xw,  sheet_name="CGML_TB", index=False)
    cgme_all.to_excel(xw, sheet_name="CGME_All", index=False)
    cgme_tb.to_excel(xw,  sheet_name="CGME_TB", index=False)
log.info("Excel written successfully.")
log.info("=== Pipeline done ===")
