import win32com.client
from openpyxl import load_workbook
from datetime import datetime

def load_subject_filters(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    filters = []
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        cell_value = row[0].value
        if cell_value:
            filters.append(str(cell_value).strip())
    return filters

def get_shared_inbox(smtp_address):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    for i in range(namespace.Folders.Count):
        mailbox = namespace.Folders.Item(i + 1)
        try:
            recipient = namespace.CreateRecipient(smtp_address)
            recipient.Resolve()
            shared_folder = namespace.GetSharedDefaultFolder(recipient, 6)  # 6 = Inbox
            return shared_folder
        except Exception:
            continue

    raise Exception(f"Shared inbox for '{smtp_address}' not found or not accessible.")

def search_emails_in_shared_inbox(inbox_folder, filters, after_date, target_email):
    matching_emails = []
    messages = inbox_folder.Items
    messages.Sort("[SentOn]", True)

    for msg in messages:
        try:
            if msg.SentOn < after_date:
                continue

            # Ensure email was sent to backtesting@abc.com
            recipients = [msg.Recipients.Item(i + 1).AddressEntry.Address.lower()
                          for i in range(msg.Recipients.Count)]

            if not any(target_email.lower() in r for r in recipients):
                continue

            subject = msg.Subject
            for pattern in filters:
                if pattern.lower() in subject.lower():
                    matching_emails.append((subject, msg.SentOn.strftime("%Y-%m-%d %H:%M")))
                    break
        except Exception:
            continue

    return matching_emails

# === SETTINGS ===
excel_path = "C:\\Path\\To\\filters.xlsx"
shared_mailbox = "backtesting@abc.com"
after_date = datetime(2025, 1, 1)
filters = load_subject_filters(excel_path)

# === RUN ===
try:
    inbox = get_shared_inbox(shared_mailbox)
    results = search_emails_in_shared_inbox(inbox, filters, after_date, shared_mailbox)

    print(f"\nFound {len(results)} matching emails:\n")
    for subject, sent_on in results:
        print(f"{sent_on} | {subject}")
except Exception as e:
    print(f"Error: {e}")
