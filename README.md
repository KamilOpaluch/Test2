import win32com.client
from openpyxl import load_workbook
from datetime import datetime

# Load search strings from Excel (column A from A2 downward)
def load_subject_filters(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    filters = []
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        cell_value = row[0].value
        if cell_value:
            filters.append(str(cell_value).strip())
    return filters

# Connect to Outlook and search emails
def search_emails(folder_name, filters, after_date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Find the 'Backtesting' folder (you may need to adjust if it's a group mailbox)
    folder = None
    for i in range(namespace.Folders.Count):
        root_folder = namespace.Folders.Item(i + 1)
        try:
            folder = root_folder.Folders(folder_name)
            break
        except:
            continue

    if folder is None:
        raise Exception(f"Folder '{folder_name}' not found.")

    matching_emails = []

    # Loop through emails
    messages = folder.Items
    messages.Sort("[SentOn]", True)  # Newest first

    for msg in messages:
        try:
            sent_on = msg.SentOn
            if sent_on < after_date:
                continue

            subject = msg.Subject
            for pattern in filters:
                if pattern.lower() in subject.lower():
                    matching_emails.append((subject, sent_on.strftime("%Y-%m-%d %H:%M")))
                    break
        except Exception as e:
            continue  # Skip broken items (e.g. meeting invites)

    return matching_emails

# === RUNNING THE SCRIPT ===
excel_path = "C:\\Path\\To\\Your\\filters.xlsx"
filters = load_subject_filters(excel_path)

after_date = datetime(2025, 1, 1)
folder_name = "Backtesting"

results = search_emails(folder_name, filters, after_date)

# Print results
print(f"\nFound {len(results)} matching emails:\n")
for subject, sent_on in results:
    print(f"{sent_on} | {subject}")
