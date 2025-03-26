import win32com.client
from datetime import datetime, timezone
import os
from openpyxl import Workbook

# === CONFIGURATION ===
shared_mailbox = "backtesting@abc.com"
after_date = datetime(2025, 1, 1, tzinfo=timezone.utc)

# Keywords that must be in the subject
contains_keywords = ["Backtesting", "VaR"]
contains_logic = "AND"  # OR or AND

# Keywords that must NOT be in the subject
not_contains_keywords = ["test", "old"]
not_contains_logic = "OR"  # OR or AND

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

def subject_matches(subject, must_have, must_not_have, logic_must="AND", logic_not="OR"):
    subj = subject.lower()

    # Handle must-have list
    if must_have:
        if logic_must == "AND":
            if not all(kw.lower() in subj for kw in must_have if kw.strip()):
                return False
        else:  # OR
            if not any(kw.lower() in subj for kw in must_have if kw.strip()):
                return False

    # Handle must-NOT-have list
    if must_not_have:
        if logic_not == "AND":
            if all(kw.lower() in subj for kw in must_not_have if kw.strip()):
                return False
        else:  # OR
            if any(kw.lower() in subj for kw in must_not_have if kw.strip()):
                return False

    return True

def search_matching_emails(inbox_folder, must_have, must_not_have, logic_must, logic_not, after_date):
    messages = inbox_folder.Items
    messages.Sort("[ReceivedTime]", True)

    results = []

    print("Scanning emails...")
    for msg in messages:
        try:
            received = msg.ReceivedTime
            subject = msg.Subject

            if received.tzinfo is None:
                received = received.replace(tzinfo=timezone.utc)

            if received < after_date:
                break

            print(f"Checking: {received.strftime('%Y-%m-%d %H:%M')} | {subject}")

            if subject_matches(subject, must_have, must_not_have, logic_must, logic_not):
                print(f"Matched: {subject}")
                results.append((subject, received.strftime("%Y-%m-%d %H:%M")))

        except Exception as e:
            print(f"Skipped due to error: {e}")
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

# === RUN ===
try:
    inbox = get_shared_inbox(shared_mailbox)
    matches = search_matching_emails(
        inbox,
        contains_keywords,
        not_contains_keywords,
        contains_logic,
        not_contains_logic,
        after_date
    )

    output_dir = os.path.dirname(os.path.abspath(__file__))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = os.path.join(output_dir, f"Filtered_Emails_{timestamp}.xlsx")

    write_to_excel(matches, output_file)

    print(f"\nSaved {len(matches)} emails to:\n{output_file}")

except Exception as e:
    print(f"Fatal Error: {e}")
