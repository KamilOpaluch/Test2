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

pattern = re.compile(r"^PROXY_(.+?)_(.+?)_(.+?)_(\d{8})\.xlsx$")
grouped_files = defaultdict(list)
unmatched_files = []

# Gather files
for filename in os.listdir(base_folder):
    if filename.endswith('.xlsx') and filename.startswith('PROXY_'):
        match = pattern.match(filename)
        if match:
            name, typ, typ2, date_str = match.groups()
            key = (name, typ, typ2)
            file_path = os.path.join(base_folder, filename)
            grouped_files[key].append((file_path, date_str))

# Process groups
for (name, typ, typ2), files in grouped_files.items():
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
        df_combined = pd.concat(all_data, ignore_index=True).iloc[:, :12]  # Keep A–L only

        output_filename = f"APPENDED_{name}_{typ}_{typ2}.xlsx"
        output_path = os.path.join(output_folder, output_filename)

        if not os.path.exists(template_path):
            print(f"Template not found: {template_path}")
            continue

        # Copy template
        shutil.copy(template_path, output_path)

        try:
            # Load template and write only data using pandas writer
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_combined.to_excel(writer, sheet_name=writer.book.sheetnames[0], startrow=0, startcol=0, index=False, header=True)
        except Exception as e:
            print(f"Error writing to file: {output_filename}: {e}")

# Report skipped
if unmatched_files:
    print("\nSkipped the following files due to no matching pair:")
    for f in unmatched_files:
        print(f"- {os.path.basename(f)}")

print("\nFast processing complete. Files saved with template preserved.")
