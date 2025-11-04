import re
from pathlib import Path
import pandas as pd

# --- Settings ---
txt_path = Path(r"C:\Users\ko59623\Documents\Test_py\email.txt")
out_xlsx = txt_path.with_name("parsed_emails.xlsx")

# --- Read file ---
# Try common encodings; fall back to errors='ignore'
for enc in ("utf-8-sig", "utf-8", "cp1250", "latin-1"):
    try:
        text = txt_path.read_text(encoding=enc)
        break
    except Exception:
        text = None
if text is None:
    raise FileNotFoundError(f"Could not read {txt_path} with tried encodings.")

# --- Find all emails that start with '<' and end with '.com' ---
# Matches like: <john.doe@citi.com> , <name.last@something.com>
email_pat = re.compile(r"<([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.com)>")

citi_soeids = []
dl_rows = []

for m in email_pat.finditer(text):
    full_email = m.group(1)  # e.g., john.doe@citi.com
    local, _, domain = full_email.partition("@")

    # Determine "name" by taking text before '<' and after last ';'
    # Slice the string up to the '<' that starts this match
    pre = text[:m.start()]
    # Find the last ';' before this '<'
    semi_idx = pre.rfind(";")
    raw_name = pre[semi_idx+1:] if semi_idx != -1 else pre
    name = raw_name.strip().strip("\"'“”‘’,.- ").replace("\n", " ").replace("\r", " ")
    # Collapse multiple spaces
    name = re.sub(r"\s{2,}", " ", name)

    if domain.lower() == "citi.com":
        # Put SOEID (local part) to SOEID sheet
        citi_soeids.append(local)
    else:
        # Put non-citi.com to DL sheet with name and full email
        dl_rows.append({"name": name, "email": full_email})

# --- Build DataFrames ---
soeid_df = pd.DataFrame({"soeid": citi_soeids})
dl_df = pd.DataFrame(dl_rows, columns=["name", "email"])

# Optional: drop exact duplicates (keep order)
if not soeid_df.empty:
    soeid_df = soeid_df.drop_duplicates().reset_index(drop=True)
if not dl_df.empty:
    dl_df = dl_df.drop_duplicates().reset_index(drop=True)

# --- Write Excel ---
with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
    # Ensure exact sheet names & column order
    soeid_df.to_excel(xw, sheet_name="SOEID", index=False)
    dl_df.to_excel(xw, sheet_name="DL", index=False)

print(f"Done. Wrote: {out_xlsx}")
