import pandas as pd
import re
from pathlib import Path

# -----------------------
# Config / Paths
# -----------------------
base = Path.home() / "Documents" / "IMA_Extend"
monthly_pattern = re.compile(r".*_(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.xlsx$", re.IGNORECASE)

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
    """Trim and collapse inner spaces in column names."""
    return [re.sub(r"\s+", " ", str(c)).strip() for c in cols]

def normalize_account_series(s: pd.Series) -> pd.Series:
    """Firm Account as string; trim; strip trailing .0 from Excel; preserve leading zeros."""
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
    # Row 12 has headers -> header index 11
    df = pd.read_excel(path, header=11, engine="openpyxl", dtype={"LV ID": str})
    df.columns = normalize_cols(df.columns)

    # Filter LV ID == "06500" (if available)
    if "LV ID" in df.columns:
        lv = (
            df["LV ID"].astype(str).str.strip()
              .str.replace(r"\.0$", "", regex=True)
              .str.zfill(5)
        )
        df = df[lv == "06500"]

    # Harmonize columns
    # Handle possible variants like leading spaces that were normalized already
    # Create empty cols if missing
    for c in WANTED:
        if c not in df.columns:
            df[c] = pd.NA

    out = df[WANTED].copy()
    out["Firm Account"] = normalize_account_series(out["Firm Account"])
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = ensure_numeric(out, ["Tot Clean P&L ex Theta", "Actual"])
    # Normalize Rwa Type values
    out["Rwa Type"] = out["Rwa Type"].astype(str).str.strip()
    return out.dropna(subset=["Date", "Firm Account"])

def load_monthlies(folder: Path) -> pd.DataFrame:
    files = [p for p in folder.glob("*.xlsx") if monthly_pattern.match(p.name)]
    if not files:
        raise FileNotFoundError(f"No monthly *_Mon.xlsx files found in {folder}")
    frames = [load_monthly_file(p) for p in files]
    if not frames:
        return pd.DataFrame(columns=WANTED)
    return pd.concat(frames, ignore_index=True)

def load_mtd_adjustments(folder: Path) -> pd.DataFrame:
    # Exact filename per instruction
    path = folder / "MTD_adjusmtents.xlsx"
    df = pd.read_excel(path, sheet_name="Adjustment", engine="openpyxl")
    df.columns = normalize_cols(df.columns)

    # Optional LV ID filter if present (align to 06500)
    if "LV ID" in df.columns:
        lv = (
            df["LV ID"].astype(str).str.strip()
              .str.replace(r"\.0$", "", regex=True)
              .str.zfill(5)
        )
        df = df[lv == "06500"]

    # Ensure all wanted columns exist
    for c in WANTED:
        if c not in df.columns:
            df[c] = pd.NA

    out = df[WANTED].copy()
    out["Firm Account"] = normalize_account_series(out["Firm Account"])
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = ensure_numeric(out, ["Tot Clean P&L ex Theta", "Actual"])
    out["Rwa Type"] = out["Rwa Type"].astype(str).str.strip()
    return out.dropna(subset=["Firm Account"])  # date may be NaT in some adjustments; we keep rows with missing date? Prefer having Date; filter if you must

def load_mcc_accounts(folder: Path) -> pd.DataFrame:
    path = folder / "mcc.xlsx"

    def load_sheet(name):
        df = pd.read_excel(path, sheet_name=name, engine="openpyxl")
        df.columns = normalize_cols(df.columns)
        # Expect LEVEL_12, LEGAL_ENTITY
        need = ["LEVEL_12", "LEGAL_ENTITY"]
        for c in need:
            if c not in df.columns:
                df[c] = pd.NA
        out = df[need].copy()
        # Remove S2 prefix from accounts, keep as string, preserve leading zeros after prefix
        out["Firm Account"] = (
            out["LEVEL_12"].astype(str).str.strip()
            .str.replace(r"^S2", "", regex=True)
            .str.replace(r"\.0$", "", regex=True)
        )
        out["LEGAL_ENTITY"] = out["LEGAL_ENTITY"].astype(str).str.strip().str.upper()
        return out[["Firm Account", "LEGAL_ENTITY"]].dropna(subset=["Firm Account"])

    new_df = load_sheet("new")
    old_df = load_sheet("old")

    # Prefer "new" mapping if duplicates
    # Combine and drop duplicates keeping first occurrence (new first)
    combo = pd.concat([new_df, old_df], ignore_index=True)
    combo = combo.drop_duplicates(subset=["Firm Account"], keep="first")
    # Keep only CGML / CGME rows if others exist
    combo = combo[combo["LEGAL_ENTITY"].isin(["CGML", "CGME"])].copy()
    return combo.reset_index(drop=True)

# -----------------------
# Orchestration
# -----------------------
# Load pieces
daily_df = load_monthlies(base)
mtd_df = load_mtd_adjustments(base)
mcc_map = load_mcc_accounts(base)

# Filter to MCC accounts (old ∪ new) and attach LEGAL_ENTITY
daily_df = daily_df.merge(mcc_map, on="Firm Account", how="inner")
mtd_df   = mtd_df.merge(mcc_map,   on="Firm Account", how="inner")

# Combine daily P&L and adjustments as additional rows
all_df = pd.concat([daily_df, mtd_df], ignore_index=True)

# Final cleanup / ordering
all_df = all_df[WANTED + ["LEGAL_ENTITY"]]
all_df = all_df.sort_values(["LEGAL_ENTITY", "Date", "Firm Account"], kind="stable").reset_index(drop=True)

# -----------------------
# Build time-series summaries
# -----------------------
def timeseries(df: pd.DataFrame, entity: str, tb_only: bool) -> pd.DataFrame:
    sub = df[df["LEGAL_ENTITY"] == entity].copy()
    if tb_only:
        # 'TB' match case-insensitive and after trimming
        sub = sub[sub["Rwa Type"].astype(str).str.strip().str.upper() == "TB"]
    # Sum by date across all accounts
    agg = (
        sub.groupby("Date", as_index=False)[["Tot Clean P&L ex Theta", "Actual"]]
           .sum()
           .sort_values("Date", kind="stable")
    )
    return agg

cgml_all = timeseries(all_df, "CGML", tb_only=False)
cgml_tb  = timeseries(all_df, "CGML", tb_only=True)
cgme_all = timeseries(all_df, "CGME", tb_only=False)
cgme_tb  = timeseries(all_df, "CGME", tb_only=True)

# -----------------------
# Write Excel output
# -----------------------
out_path = base / "pnl_timeseries_summary.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
    # Raw combined rows (optional but handy for QA)
    all_df.to_excel(xw, sheet_name="Combined_Rows", index=False)

    # Account mapping (for traceability)
    mcc_map.to_excel(xw, sheet_name="Account_Map", index=False)

    # Time series per entity & scenario
    cgml_all.to_excel(xw, sheet_name="CGML_All", index=False)
    cgml_tb.to_excel(xw,  sheet_name="CGML_TB", index=False)
    cgme_all.to_excel(xw, sheet_name="CGME_All", index=False)
    cgme_tb.to_excel(xw,  sheet_name="CGME_TB", index=False)

print("Done.")
print(f"Daily rows (race clean): {len(daily_df)}")
print(f"Adjustment rows: {len(mtd_df)}")
print(f"Output written to: {out_path}")
