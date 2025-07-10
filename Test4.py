import os
import re
import shutil
import pandas as pd
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook

base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'QE_Proxy')
output_folder = os.path.join(base_folder, 'Processed')
template_path = os.path.join(output_folder, 'Template.xlsx')
os.makedirs(output_folder, exist_ok=True)

# Patterns:
# 1. With type2: PROXY_name_type1_type2_yyyymmdd.xlsx
# 2. Without type2: PROXY_name_type1_yyyymmdd.xlsx
pattern_with_type2 = re.compile(r"^PROXY_(.+?)_(.+?)_(.+?)_(\d{8})\.xlsx$")
pattern_no_type2 = re.compile(r"^PROXY_(.+?)_(.+?)_(\d{8})\.xlsx$")

grouped_files = defaultdict(list)
unmatched_files = []

# Gather files
for filename in os.listdir(base_folder):
    if not filename.endswith('.xlsx') or not filename.startswith('PROXY_'):
        continue

    full_path = os.path.join(base_folder, filename)

    match = pattern_with_type2.match(filename)
    if match:
        name, type1, type2, date_str = match.groups()
        key = (name, type1, type2)
    else:
        match = pattern_no_type2.match(filename)
        if match:
            name, type1, date_str = match.groups()
            key = (name, type1, None)
        else:
            print(f"⚠️ Filename skipped (no match): {filename}")
            continue

    grouped_files[key].append((full_path, date_str))

# Process groups
for (name, type1, type2), files in grouped_files.items():
    if len(files) < 2:
        unmatched_files.append(files[0][0])
        continue

    all_data = []
    for file_path, date_str in files:
        try:
            df = pd.read_excel(file_path, skiprows=2)
            formatted_date = datetime.strptime(date_str, "%Y%m%d").strftime("%m/%d/%Y")
            df.insert(0, 'Date', formatted_date)
            all_data.append(df)
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    if all_data:
        df_combined = pd.concat(all_data, ignore_index=True).iloc[:, :12]  # A–L only

        # Build output filename
        typ2_part = f"_{type2}" if type2 else ""
        output_filename = f"APPENDED_{name}_{type1}{typ2_part}.xlsx"
        output_path = os.path.join(output_folder, output_filename)

        if not os.path.exists(template_path):
            print(f"❌ Template.xlsx missing in: {template_path}")
            continue

        shutil.copy(template_path, output_path)

        try:
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_combined.to_excel(
                    writer,
                    sheet_name=writer.book.sheetnames[0],
                    startrow=0,
                    startcol=0,
                    index=False,
                    header=True
                )
        except Exception as e:
            print(f"Error writing to file: {output_filename}: {e}")

# Report skipped
if unmatched_files:
    print("\n📄 Skipped files with no matching pair:")
    for f in unmatched_files:
        print(f"- {os.path.basename(f)}")

print("\n✅ Processing complete. All matched groups saved with template.")
