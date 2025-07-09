import os
import re
import pandas as pd
from datetime import datetime
from collections import defaultdict

# Define folders
base_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'QE_Proxy')
output_folder = os.path.join(base_folder, 'Processed')
os.makedirs(output_folder, exist_ok=True)

# Regex pattern to match filenames
pattern = re.compile(r"^PROXY_(.+?)_(.+?)_(.+?)_(\d{8})\.xlsx$")

# Store grouped data
grouped_files = defaultdict(list)
unmatched_files = []

# Walk through the folder and collect files
for filename in os.listdir(base_folder):
    if filename.endswith('.xlsx') and filename.startswith('PROXY_'):
        match = pattern.match(filename)
        if match:
            name, typ, typ2, date_str = match.groups()
            group_key = (name, typ, typ2)
            file_path = os.path.join(base_folder, filename)
            grouped_files[group_key].append((file_path, date_str))

# Process each group
for (name, typ, typ2), files in grouped_files.items():
    if len(files) < 2:
        unmatched_files.append(files[0][0])  # only one file in group
        continue

    all_data = []
    for file_path, date_str in files:
        try:
            df = pd.read_excel(file_path)
            date_formatted = datetime.strptime(date_str, "%Y%m%d").strftime("%m/%d/%Y")
            df.insert(0, 'Date', date_formatted)
            all_data.append(df)
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        output_file = os.path.join(output_folder, f"APPENDED_{name}_{typ}_{typ2}.xlsx")
        combined_df.to_excel(output_file, index=False)

# Report unmatched files
if unmatched_files:
    print("\nSkipped the following files due to no matching pair:")
    for f in unmatched_files:
        print(f"- {os.path.basename(f)}")
else:
    print("\nAll files were matched and processed.")

print("\nProcessing complete. Appended files saved in 'Processed' subfolder.")
