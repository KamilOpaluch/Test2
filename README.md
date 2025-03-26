import win32com.client
from datetime import datetime
import os
from openpyxl import Workbook

# === CONFIGURATION ===
shared_mailbox = "backtesting@abc.com"  # Can also be just "Backtesting"
subject_keyword = "Backtesting VaR"
after_date = datetime(2025, 1, 1)

def get_shared_inbox(smtp_address_or_display_name):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    recipient = namespace.CreateRecipient(smtp_address_or_display_name)
    recipient.Resolve()
    if recipient.Resolved:
        print(f"Connected to inbox of: {smtp_address_or_display_name}")
        inbox = namespace.GetSharedDefaultFolder(recipient, 6)  # 6 = Inbox
        return inbox
    else:
        raise Exception(f"Could not resolve shared mailbox: {smtp_address_or_display_name}")

def recipient_matches(msg, target_email_or_name):
    target = target_email_or_name.lower()
    for i in range(msg.Recipients.Count):
        r = msg.Recipients.Item(i + 1)
        try:
            name = r.Name.lower()
            address = r.AddressEntry.Address.lower()
            if target in name or target in address:
                return True
        except:
            continue
    return False

def search_matching_emails(inbox_folder, keyword, after_date, target_email_or_name):
    messages = inbox_folder.Items

    # Use Restrict to filter by date and subject
    restriction = f"[ReceivedTime] >= '{after_date.strftime('%m/%d/%Y %I:%M %p')}' AND [Subject] LIKE '%{keyword}%'"
    filtered = messages.Restrict(restriction)
    filtered.Sort("[ReceivedTime]", True)

    results = []

    for msg in filtered:
        try:
            subject = msg.Subject
            received = msg.ReceivedTime
            sender = msg.SenderEmailAddress

            print(f"Checking: {received.strftime('%Y-%m-%d %H:%M')} | {subject} | From: {sender}")

            if not recipient_matches(msg, target_email_or_name):
                print(f"Skipped (recipient not matching): {subject}")
                continue

            print(f"Matched: {subject}")
            results.append((subject, received.strftime("%Y-%m-%d %H:%M")))

        except Exception as e:
            print(f"Error processing message: {e}")
            continue

    return results

def write_to_excel(data, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Matching Emails"
    ws.append(["Subject", "Received Date"])
    for subject, received in data:
        ws.append([subject, received])
    wb.save(output_path)

# === RUN SCRIPT ===
try:
    inbox = get_shared_inbox(shared_mailbox)
    matches = search_matching_emails(inbox, subject_keyword, after_date, shared_mailbox)

    # Save to Excel in script's folder
    output_dir = os.path.dirname(os.path.abspath(__file__))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = os.path.join(output_dir, f"Backtesting_VAR_Emails_{timestamp}.xlsx")

    write_to_excel(matches, output_file)

    print(f"\nSaved {len(matches)} matching emails to:\n{output_file}")

except Exception as e:
    print(f"Error: {e}")
