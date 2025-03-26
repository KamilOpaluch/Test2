import win32com.client
from datetime import datetime
import os
from openpyxl import Workbook

# === CONFIGURATION ===
shared_mailbox = "backtesting@abc.com"
subject_keyword = "Backtesting VaR"
after_date = datetime(2025, 1, 1)

def get_shared_inbox(smtp_address):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    recipient = namespace.CreateRecipient(smtp_address)
    recipient.Resolve()
    if recipient.Resolved:
        inbox = namespace.GetSharedDefaultFolder(recipient, 6)  # 6 = Inbox
        return inbox
    else:
        raise Exception(f"Could not resolve shared mailbox: {smtp_address}")

def search_matching_emails(inbox_folder, keyword, after_date, target_email):
    messages = inbox_folder.Items
    messages.Sort("[SentOn]", True)

    results = []

    for msg in messages:
        try:
            if msg.SentOn < after_date:
                continue

            # Check if sent to the shared mailbox
            recipients = [msg.Recipients.Item(i + 1).AddressEntry.Address.lower()
                          for i in range(msg.Recipients.Count)]
            if not any(target_email.lower() in r for r in recipients):
                continue

            if keyword.lower() in msg.Subject.lower():
                results.append((msg.Subject, msg.SentOn.strftime("%Y-%m-%d %H:%M")))
        except Exception:
            continue

    return results

def write_to_excel(data, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Matching Emails"
    ws.append(["Subject", "Sent Date"])
    for subject, sent_on in data:
        ws.append([subject, sent_on])
    wb.save(output_path)

# === RUN ===
try:
    inbox = get_shared_inbox(shared_mailbox)
    matches = search_matching_emails(inbox, subject_keyword, after_date, shared_mailbox)

    # Save Excel in same folder as script
    output_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(output_dir, f"Backtesting_VAR_Emails_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    write_to_excel(matches, output_file)

    print(f"Saved {len(matches)} emails to:\n{output_file}")

except Exception as e:
    print(f"Error: {e}")
