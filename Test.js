import pandas as pd
import re
import sys
import logging
from pathlib import Path

# -----------------------
# Logging
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
WANTED  = ["MSH Level 4","Firm Account","Date","Tot Clean P&L ex Theta","Actual","Rwa Type"]

# Allow both LVs and map them to entity
ALLOWED_LVIDS = {"06500": "CGML", "02937": "CGME"}
ALLOWED_LVSET = set(ALLOWED_LVIDS.keys())

# -----------------------
# Helpers
# -----------------------
def normalize_cols(cols): return [re.sub(r"\s+", " ", str(c)).strip() for c in cols]

def normalize_account_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

def make_acc_key(s: pd.Series) -> pd.Series:
    key = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    key = key.str.lstrip("0")
    return key.mask(key.eq(""), "0")

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
def _postprocess_common(df: pd.DataFrame) -> pd.DataFrame:
    # Ensure wanted cols
    for c in WANTED:
        if c not in df.columns:
            df[c] = pd.NA

    out = df[WANTED].copy()
    out["Firm Account"] = normalize_account_series(out["Firm Account"])
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = ensure_numeric(out, ["Tot Clean P&L ex Theta","Actual"])
    out["Rwa Type"] = out["Rwa Type"].astype(str).str.strip()

    # Keys & hints
    out["acc_key"]     = make_acc_key(out["Firm Account"])
    out["entity_hint"] = entity_hint_from_msh4(out["MSH Level 4"])
    return out

def load_monthly_file(path: Path) -> pd.DataFrame:
    try:
        log.info("Reading monthly file: %s", path.name)
        df = pd.read_excel(path, header=11, engine="openpyxl", usecols=USECOLS, dtype={"LV ID": str})
        df.columns = normalize_cols(df.columns)
        log.info("[%s] Loaded %d rows", path.name, len(df))

        # Normalize LV and keep BOTH 06500/02937
        if "LV ID" in df.columns:
            lv = (df["LV ID"].astype(str).str.strip()
                               .str.replace(r"\.0$", "", regex=True)
                               .str.zfill(5))
            df["LV_norm"] = lv
            before = len(df)
            df = df[lv.isin(ALLOWED_LVSET)]
            log.info("[%s] LV filter kept %d / %d rows (06500=CGML, 02937=CGME)", path.name, len(df), before)
            # Map LV → entity (strongest signal)
            df["lv_entity"] = df["LV_norm"].map(ALLOWED_LVIDS)
        else:
            log.warning("[%s] No 'LV ID' column; keeping all rows and lv_entity=NaN", path.name)
            df["LV_norm"]  = pd.NA
            df["lv_entity"] = pd.NA

        out = _postprocess_common(df)
        # carry lv_entity through
        out["lv_entity"] = df["lv_entity"].reset_index(drop=True)

        # Drop rows without Date or Account
        before = len(out)
        out = out.dropna(subset=["Date","Firm Account"])
        log.info("[%s] After normalize & non-null (Date/Account): %d rows (dropped %d)", path.name, len(out), before - len(out))
        return out
    except Exception as e:
        log.exception("Failed reading %s: %s", path.name, e)
        cols = WANTED + ["acc_key","entity_hint","lv_entity"]
        return pd.DataFrame(columns=cols)

def load_monthlies(folder: Path) -> pd.DataFrame:
    files = [p for p in folder.glob("*.xlsx") if monthly_pattern.match(p.name)]
    log.info("Discovered %d monthly files in %s", len(files), folder)
    if not files:
        raise FileNotFoundError(f"No monthly *_Mon.xlsx files found in {folder}")
    frames = [load_monthly_file(p) for p in files]
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=WANTED + ["acc_key","entity_hint","lv_entity"])
    log.info("Monthly combined: %d rows, %d unique accounts", len(combined), combined["Firm Account"].nunique())
    # LV distribution diag
    if "lv_entity" in combined.columns:
        log.info("LV-derived entity distribution:\n%s", combined["lv_entity"].value_counts(dropna=False).to_string())
    return combined

def load_mtd_adjustments(folder: Path) -> pd.DataFrame:
    path = folder / "MTD_adjustment.xlsx"
    log.info("Reading MTD adjustments: %s (sheet='Adjustment')", path.name)
    try:
        df = pd.read_excel(path, sheet_name="Adjustment", engine="openpyxl")
    except Exception as e:
        log.exception("Failed to read %s: %s", path, e)
        return pd.DataFrame(columns=WANTED + ["acc_key","entity_hint","lv_entity"])

    df.columns = normalize_cols(df.columns)
    # Optional LV filter if present (keep both CGML & CGME)
    if "LV ID" in df.columns:
        lv = (df["LV ID"].astype(str).str.strip()
                           .str.replace(r"\.0$", "", regex=True)
                           .str.zfill(5))
        df["LV_norm"] = lv
        before = len(df)
        df = df[lv.isin(ALLOWED_LVSET)]
        log.info("[MTD] LV filter kept %d / %d rows", len(df), before)
        df["lv_entity"] = df["LV_norm"].map(ALLOWED_LVIDS)
    else:
        df["lv_entity"] = pd.NA

    out = _postprocess_common(df)
    out["lv_entity"] = df["lv_entity"].reset_index(drop=True)

    kept = len(out.dropna(subset=["Firm Account"]))
    log.info("[MTD] Prepared rows: %d (Firm Account present), unique accounts: %d", kept, out["Firm Account"].nunique())
    return out.dropna(subset=["Firm Account"])

def load_mcc_accounts(folder: Path) -> pd.DataFrame:
    path = folder / "mcc.xlsx"
    log.info("Reading account map: %s ('new' & 'old')", path.name)

    def load_sheet(name, src):
        try:
            df = pd.read_excel(path, sheet_name=name, engine="openpyxl")
        except Exception as e:
            log.exception("Failed reading %s sheet '%s': %s", path.name, name, e)
            return pd.DataFrame(columns=["Firm Account","acc_key","LEGAL_ENTITY","source"])
        df.columns = normalize_cols(df.columns)
        for c in ["LEVEL_12","LEGAL_ENTITY"]:
            if c not in df.columns: df[c] = pd.NA
        out = pd.DataFrame({
            "Firm Account": (df["LEVEL_12"].astype(str).str.strip()
                                               .str.replace(r"^S2[\s\-_]*", "", regex=True)
                                               .str.replace(r"\.0$", "", regex=True)),
            "LEGAL_ENTITY": df["LEGAL_ENTITY"].astype(str).str.strip().str.upper(),
            "source": src
        })
        out["acc_key"] = make_acc_key(out["Firm Account"])
        out = out[["Firm Account","acc_key","LEGAL_ENTITY","source"]]
        log.info("[mcc:%s] rows=%d, unique accounts=%d", name, len(out), out["Firm Account"].nunique())
        return out

    m_new = load_sheet("new","new")
    m_old = load_sheet("old","old")
    mcc = pd.concat([m_new,m_old], ignore_index=True)
    mcc = mcc[mcc["LEGAL_ENTITY"].isin(["CGML","CGME"])].copy()  # keep both; duplicates allowed
    log.info("[mcc] Combined rows=%d | unique accounts=%d | CGML rows=%d | CGME rows=%d",
             len(mcc), mcc["Firm Account"].nunique(),
             (mcc["LEGAL_ENTITY"]=="CGML").sum(),
             (mcc["LEGAL_ENTITY"]=="CGME").sum())
    return mcc.reset_index(drop=True)

# -----------------------
# Entity assignment & conflict resolution
# -----------------------
def attach_entity_with_resolution(daily_like: pd.DataFrame, mcc_map: pd.DataFrame) -> pd.DataFrame:
    """
    Assign a single LEGAL_ENTITY per row using priority:
    1) LV-derived entity (lv_entity) if present on the row AND exists in MCC candidates
    2) MSH Level 4 hint match, if exists in MCC candidates
    3) Prefer MCC source 'new'
    4) Fallback to any mapping
    """
    if daily_like.empty:
        daily_like["LEGAL_ENTITY"] = pd.NA
        return daily_like

    work = daily_like.reset_index(drop=False).rename(columns={"index":"_row_id"})
    merged = work.merge(mcc_map, on="acc_key", how="left", suffixes=("","_map"))

    # Flags against each candidate
    merged["lv_match"]   = merged["lv_entity"].notna() & merged["LEGAL_ENTITY"].eq(merged["lv_entity"])
    merged["hint_match"] = merged["entity_hint"].notna() & merged["LEGAL_ENTITY"].eq(merged["entity_hint"])
    merged["source_rank"] = merged["source"].map({"new":0, "old":1}).fillna(2).astype(int)

    # Sort candidates per row by our priority
    merged = merged.sort_values(
        ["_row_id", "lv_match", "hint_match", "source_rank"],
        ascending=[True, False, False, True]
    )

    picked = merged.drop_duplicates(subset=["_row_id"], keep="first")
    total = len(work)
    log.info("Entity resolution: rows=%d | lv-match=%d | hint-match(no lv)=%d | prefer-new=%d | fallback=%d",
             total,
             picked["lv_match"].sum(),
             ((picked["lv_match"]==False) & (picked["hint_match"]==True)).sum(),
             ((picked["lv_match"]==False) & (picked["hint_match"]==False) & (picked["source_rank"]==0)).sum(),
             ((picked["lv_match"]==False) & (picked["hint_match"]==False) & (picked["source_rank"]>0)).sum()
    )

    out = picked.drop(columns=["lv_match","hint_match","source_rank","source"])
    # rows with no mapping at all
    missing = out["LEGAL_ENTITY"].isna().sum()
    if missing:
        log.warning("Rows without any MCC mapping: %d (they will be dropped next)", missing)
    return out

def enforce_one_entity_per_day(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure at most one LEGAL_ENTITY per (Date, Firm Account).
    Priority inside a day:
      1) rows whose LEGAL_ENTITY equals lv_entity (if any)
      2) entity with most rows that day
      3) alphabetical tie-break
    We *keep all rows* that match the chosen entity for that day/account (so multiple lines for the chosen entity are preserved).
    Rows mapped to the non-chosen entity for that same day/account are dropped.
    """
    if df.empty:
        return df

    # Count entity rows per (day, account)
    cnt = (df.groupby(["Date","Firm Account","LEGAL_ENTITY"])
             .size().rename("cnt").reset_index())
    # Get group-level lv (first non-null)
    grp_lv = (df.groupby(["Date","Firm Account"])["lv_entity"]
                .apply(lambda s: next((x for x in s.dropna().unique()), pd.NA))
                .reset_index().rename(columns={"lv_entity":"grp_lv"}))

    choice = cnt.merge(grp_lv, on=["Date","Firm Account"], how="left")
    choice["prefer_lv"] = choice["grp_lv"].notna() & choice["LEGAL_ENTITY"].eq(choice["grp_lv"])
    # Sort by: prefer_lv desc, cnt desc, LEGAL_ENTITY asc
    choice = choice.sort_values(["Date","Firm Account","prefer_lv","cnt","LEGAL_ENTITY"],
                                ascending=[True, True, False, False, True])
    choice = choice.drop_duplicates(["Date","Firm Account"], keep="first")
    # Now keep only rows whose entity equals the chosen entity for that day/account
    before = len(df)
    df2 = df.merge(choice[["Date","Firm Account","LEGAL_ENTITY"]], on=["Date","Firm Account","LEGAL_ENTITY"], how="inner")
    conflicts = (cnt.groupby(["Date","Firm Account"]).size() > 1).sum()
    log.info("Per-day entity conflicts resolved: %d groups | kept %d / %d rows", conflicts, len(df2), before)
    return df2

# -----------------------
# Time series
# -----------------------
def timeseries(df: pd.DataFrame, entity: str, tb_only: bool) -> pd.DataFrame:
    sub = df[df["LEGAL_ENTITY"] == entity].copy()
    if tb_only:
        sub = sub[sub["Rwa Type"].astype(str).str.strip().str.upper() == "TB"]
    if sub.empty:
        log.warning("[TS] %s | TB=%s -> no rows", entity, tb_only)
        return pd.DataFrame(columns=["Date","Tot Clean P&L ex Theta","Actual"])
    agg = (sub.groupby("Date", as_index=False)[["Tot Clean P&L ex Theta","Actual"]]
             .sum()
             .sort_values("Date", kind="stable"))
    dmin, dmax = sub["Date"].min(), sub["Date"].max()
    log.info("[TS] %s | TB=%s -> %d dates (%s .. %s)", entity, tb_only, agg.shape[0],
             dmin.date() if pd.notna(dmin) else "NaT",
             dmax.date() if pd.notna(dmax) else "NaT")
    return agg

# -----------------------
# Orchestration
# -----------------------
log.info("=== Pipeline start ===")
daily_raw = load_monthlies(base)
mtd_raw   = load_mtd_adjustments(base)
mcc_map   = load_mcc_accounts(base)

# Attach entity with LV→entity priority
daily_df = attach_entity_with_resolution(daily_raw, mcc_map)
mtd_df   = attach_entity_with_resolution(mtd_raw,   mcc_map)

# Drop non-mapped rows and non-target entities (should only be CGML/CGME anyway)
daily_df = daily_df[daily_df["LEGAL_ENTITY"].isin(["CGML","CGME"])]
mtd_df   = mtd_df[mtd_df["LEGAL_ENTITY"].isin(["CGML","CGME"])]

log.info("Daily after entity attach: %d rows | CGML=%d | CGME=%d",
         len(daily_df), (daily_df["LEGAL_ENTITY"]=="CGML").sum(), (daily_df["LEGAL_ENTITY"]=="CGME").sum())
log.info("MTD   after entity attach: %d rows | CGML=%d | CGME=%d",
         len(mtd_df), (mtd_df["LEGAL_ENTITY"]=="CGML").sum(), (mtd_df["LEGAL_ENTITY"]=="CGME").sum())

# Combine and enforce "one entity per (Date, Firm Account)"
all_df = pd.concat([daily_df, mtd_df], ignore_index=True)
all_df = enforce_one_entity_per_day(all_df)

# Order & tidy
all_df = all_df[WANTED + ["LEGAL_ENTITY","lv_entity","entity_hint","acc_key"]]
all_df = all_df.sort_values(["LEGAL_ENTITY","Date","Firm Account"], kind="stable").reset_index(drop=True)

log.info("Combined rows after enforcement: %d | unique dates=%d | accounts=%d",
         len(all_df), all_df["Date"].nunique(), all_df["Firm Account"].nunique())

# Build time series
cgml_all = timeseries(all_df, "CGML", tb_only=False)
cgml_tb  = timeseries(all_df, "CGML", tb_only=True)
cgme_all = timeseries(all_df, "CGME", tb_only=False)
cgme_tb  = timeseries(all_df, "CGME", tb_only=True)

# Write output
out_path = base / "pnl_timeseries_summary.xlsx"
log.info("Writing Excel: %s", out_path)
with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
    all_df.to_excel(xw, sheet_name="Combined_Rows", index=False)
    mcc_map.to_excel(xw, sheet_name="Account_Map", index=False)
    cgml_all.to_excel(xw, sheet_name="CGML_All", index=False)
    cgml_tb.to_excel(xw,  sheet_name="CGML_TB", index=False)
    cgme_all.to_excel(xw, sheet_name="CGME_All", index=False)
    cgme_tb.to_excel(xw,  sheet_name="CGME_TB", index=False)
log.info("Excel written.")
log.info("=== Pipeline done ===")
