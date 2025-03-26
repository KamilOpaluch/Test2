import win32com.client
import os
from datetime import datetime, timezone
from tkinter import (
    Tk, Label, Entry, Text, Button, Scrollbar, filedialog,
    END, Frame, StringVar, IntVar, Checkbutton
)
from tkinter.ttk import Combobox
from openpyxl import Workbook
from tkinter.scrolledtext import ScrolledText


def get_outlook_inbox(mailbox):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    recipient = namespace.CreateRecipient(mailbox)
    recipient.Resolve()
    if recipient.Resolved:
        inbox = namespace.GetSharedDefaultFolder(recipient, 6)  # 6 = Inbox
        return inbox
    else:
        raise Exception(f"Cannot resolve mailbox: {mailbox}")


def search_emails(filters):
    inbox = get_outlook_inbox(filters['mailbox'])
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    results = []
    count = 0

    for msg in messages:
        try:
            received = msg.ReceivedTime
            if received.tzinfo is None:
                received = received.replace(tzinfo=timezone.utc)
            if received < filters['start_date'] or received > filters['end_date']:
                continue

            subject = msg.Subject or ""
            body = msg.Body or ""
            sender = msg.SenderEmailAddress or ""
            attachments = [att.FileName for att in msg.Attachments]

            if filters['subject_include']:
                if filters['subject_logic'] == "AND":
                    if not all(kw.lower() in subject.lower() for kw in filters['subject_include']):
                        continue
                else:
                    if not any(kw.lower() in subject.lower() for kw in filters['subject_include']):
                        continue

            if filters['subject_exclude']:
                if any(kw.lower() in subject.lower() for kw in filters['subject_exclude']):
                    continue

            if filters['body_keywords']:
                if filters['body_logic'] == "AND":
                    if not all(kw.lower() in body.lower() for kw in filters['body_keywords']):
                        continue
                else:
                    if not any(kw.lower() in body.lower() for kw in filters['body_keywords']):
                        continue

            if filters['attachment_keywords']:
                if not any(
                    any(kw.lower() in att.lower() for kw in filters['attachment_keywords'])
                    for att in attachments
                ):
                    continue

            if filters['sender_keywords']:
                if not any(kw.lower() in sender.lower() for kw in filters['sender_keywords']):
                    continue

            results.append({
                "subject": subject,
                "received": received.strftime("%Y-%m-%d %H:%M"),
                "body": body,
                "body_preview": body[:50],
                "recipients": [msg.Recipients.Item(i + 1).Address for i in range(msg.Recipients.Count)],
                "cc": [msg.CC] if msg.CC else [],
                "attachments": attachments,
                "entryid": msg.EntryID
            })

            count += 1
            if filters['limit'] and count >= filters['limit']:
                break
        except Exception:
            continue

    return results


def open_email(entryid):
    outlook = win32com.client.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")
    mail_item = session.GetItemFromID(entryid)
    mail_item.Display()


def export_to_excel(results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    headers = ["Subject", "Received", "Body", "Recipients", "CC", "Attachments"]
    ws.append(headers)
    for r in results:
        ws.append([
            r["subject"],
            r["received"],
            r["body"],
            ", ".join(r["recipients"]),
            ", ".join(r["cc"]),
            ", ".join(r["attachments"]),
        ])
    wb.save(output_path)


def run_search_ui():
    filters = {
        'mailbox': mailbox_entry.get(),
        'subject_include': subject_include.get().split(',') if subject_include.get() else [],
        'subject_exclude': subject_exclude.get().split(',') if subject_exclude.get() else [],
        'body_keywords': body_keywords.get().split(',') if body_keywords.get() else [],
        'attachment_keywords': attachment_keywords.get().split(',') if attachment_keywords.get() else [],
        'sender_keywords': sender_keywords.get().split(',') if sender_keywords.get() else [],
        'subject_logic': subject_logic.get(),
        'body_logic': body_logic.get(),
        'start_date': datetime.strptime(start_date.get(), "%Y-%m-%d").replace(tzinfo=timezone.utc),
        'end_date': datetime.strptime(end_date.get(), "%Y-%m-%d").replace(tzinfo=timezone.utc),
        'limit': int(limit.get()) if limit.get().isdigit() else None
    }

    result_output.delete("1.0", END)
    global found_emails
    found_emails = search_emails(filters)

    for idx, email in enumerate(found_emails):
        display_text = f"{idx+1}. {email['received']} | {email['subject']} | {email['body_preview']}..."
        result_output.insert(END, display_text + "\n")
        result_output.insert(END, f"[Open Email]\n")
        result_output.window_create(END, window=Button(root, text="Open", command=lambda eid=email["entryid"]: open_email(eid)))
        result_output.insert(END, "\n\n")


def export_results_ui():
    if not found_emails:
        return
    directory = filedialog.askdirectory(initialdir=os.path.expanduser("~/Documents"))
    if not directory:
        return
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filepath = os.path.join(directory, f"Filtered_Emails_{timestamp}.xlsx")
    export_to_excel(found_emails, filepath)


root = Tk()
root.title("Outlook Email Search Tool")

Label(root, text="Shared Mailbox (e.g. backtesting@abc.com):").grid(row=0, column=0, sticky="w")
mailbox_entry = Entry(root, width=50)
mailbox_entry.grid(row=0, column=1)

Label(root, text="Subject must contain (comma-separated):").grid(row=1, column=0, sticky="w")
subject_include = Entry(root, width=50)
subject_include.grid(row=1, column=1)

Label(root, text="Subject must NOT contain:").grid(row=2, column=0, sticky="w")
subject_exclude = Entry(root, width=50)
subject_exclude.grid(row=2, column=1)

Label(root, text="Body must contain:").grid(row=3, column=0, sticky="w")
body_keywords = Entry(root, width=50)
body_keywords.grid(row=3, column=1)

Label(root, text="Attachment filename contains:").grid(row=4, column=0, sticky="w")
attachment_keywords = Entry(root, width=50)
attachment_keywords.grid(row=4, column=1)

Label(root, text="Sender must contain:").grid(row=5, column=0, sticky="w")
sender_keywords = Entry(root, width=50)
sender_keywords.grid(row=5, column=1)

Label(root, text="Subject logic (AND/OR):").grid(row=6, column=0, sticky="w")
subject_logic = Combobox(root, values=["AND", "OR"])
subject_logic.set("AND")
subject_logic.grid(row=6, column=1)

Label(root, text="Body logic (AND/OR):").grid(row=7, column=0, sticky="w")
body_logic = Combobox(root, values=["AND", "OR"])
body_logic.set("OR")
body_logic.grid(row=7, column=1)

Label(root, text="Start Date (YYYY-MM-DD):").grid(row=8, column=0, sticky="w")
start_date = Entry(root, width=20)
start_date.insert(0, "2025-01-01")
start_date.grid(row=8, column=1, sticky="w")

Label(root, text="End Date (YYYY-MM-DD):").grid(row=9, column=0, sticky="w")
end_date = Entry(root, width=20)
end_date.insert(0, datetime.today().strftime("%Y-%m-%d"))
end_date.grid(row=9, column=1, sticky="w")

Label(root, text="Max emails to return (optional):").grid(row=10, column=0, sticky="w")
limit = Entry(root, width=10)
limit.grid(row=10, column=1, sticky="w")

Button(root, text="Run Search", command=run_search_ui).grid(row=11, column=0, pady=10)
Button(root, text="Export to Excel", command=export_results_ui).grid(row=11, column=1, pady=10)

Label(root, text="Results:").grid(row=12, column=0, sticky="nw")
result_output = ScrolledText(root, height=20, width=100)
result_output.grid(row=13, column=0, columnspan=2)

found_emails = []
root.mainloop()

