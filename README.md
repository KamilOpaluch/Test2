import threading
import win32com.client
import pythoncom
import os
from datetime import datetime, timezone
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Style
from openpyxl import Workbook
from tkcalendar import DateEntry
from tkinter.scrolledtext import ScrolledText

# === Outlook and Email Logic ===

def get_outlook_inbox(mailbox=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if mailbox:
        recipient = namespace.CreateRecipient(mailbox)
        recipient.Resolve()
        if recipient.Resolved:
            return namespace.GetSharedDefaultFolder(recipient, 6)
        raise Exception(f"Cannot access mailbox: {mailbox}")
    return namespace.GetDefaultFolder(6)

def match_keywords(text, keywords, logic):
    keywords = [k.strip() for k in keywords if k.strip()]
    if not keywords:
        return True
    text = text.lower()
    return all(k.lower() in text for k in keywords) if logic == "AND" else any(k.lower() in text for k in keywords)

def search_emails_thread(filters, on_result, on_done):
    pythoncom.CoInitialize()
    try:
        inbox = get_outlook_inbox(filters.get('mailbox'))
        items = inbox.Items
        restriction = "[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'".format(
            filters['start_date'].strftime("%m/%d/%Y %I:%M %p"),
            filters['end_date'].strftime("%m/%d/%Y %I:%M %p")
        )
        messages = items.Restrict(restriction)
        messages.Sort("[ReceivedTime]", True)

        count = 0
        results = []

        for msg in messages:
            try:
                subject = msg.Subject or ""
                body = msg.Body or ""
                sender = msg.SenderEmailAddress or ""
                attachments = [att.FileName for att in msg.Attachments]
                recipients = [msg.Recipients.Item(i + 1).Address for i in range(msg.Recipients.Count)]
                cc = msg.CC if msg.CC else ""
                received = msg.ReceivedTime

                if not match_keywords(subject, filters['subject_include'], filters['subject_logic']):
                    continue
                if match_keywords(subject, filters['subject_exclude'], filters['subject_exclude_logic']):
                    continue
                if not match_keywords(body, filters['body_keywords'], filters['body_logic']):
                    continue
                if not match_keywords(sender, filters['sender_keywords'], filters['sender_logic']):
                    continue
                if filters['attachment_keywords']:
                    all_attachments = ' '.join(attachments).lower()
                    if not match_keywords(all_attachments, filters['attachment_keywords'], filters['attachment_logic']):
                        continue

                result = {
                    "subject": subject,
                    "received": received.strftime("%Y-%m-%d %H:%M"),
                    "body": body,
                    "body_preview": body[:50].replace("\n", " "),
                    "recipients": ", ".join(recipients),
                    "cc": cc,
                    "attachments": ", ".join(attachments),
                    "entryid": msg.EntryID
                }

                results.append(result)
                on_result(result)

                count += 1
                if filters['limit'] and count >= filters['limit']:
                    break
            except Exception:
                continue

        on_done(results)
    except Exception as e:
        on_done([], error=str(e))

# === GUI Setup ===

def run_search():
    try:
        filters = {
            'mailbox': mailbox_entry.get().strip() or None,
            'subject_include': subject_include.get().split(','),
            'subject_exclude': subject_exclude.get().split(','),
            'body_keywords': body_keywords.get().split(','),
            'attachment_keywords': attachment_keywords.get().split(','),
            'sender_keywords': sender_keywords.get().split(','),
            'subject_logic': subject_logic.get(),
            'subject_exclude_logic': subject_exclude_logic.get(),
            'body_logic': body_logic.get(),
            'sender_logic': sender_logic.get(),
            'attachment_logic': attachment_logic.get(),
            'start_date': datetime.combine(start_date.get_date(), datetime.min.time()).replace(tzinfo=timezone.utc),
            'end_date': datetime.combine(end_date.get_date(), datetime.max.time()).replace(tzinfo=timezone.utc),
            'limit': int(limit.get()) if limit.get().isdigit() else None
        }

        output_box.delete("1.0", END)
        search_button.config(state=DISABLED)
        status_label.config(text="Searching...")

        global found_emails
        found_emails = []

        threading.Thread(
            target=search_emails_thread,
            args=(filters, on_result_found, on_search_done),
            daemon=True
        ).start()

    except Exception as e:
        messagebox.showerror("Error", str(e))
        search_button.config(state=NORMAL)

def on_result_found(result):
    found_emails.append(result)
    output_box.insert(END, f"{len(found_emails)}. {result['received']} | {result['subject']} | {result['body_preview']}\n")
    output_box.see(END)

def on_search_done(results, error=None):
    search_button.config(state=NORMAL)
    if error:
        messagebox.showerror("Search Error", error)
        status_label.config(text="Search failed.")
    else:
        if not results:
            output_box.insert(END, "No matching emails found.\n")
        status_label.config(text=f"Search complete. {len(results)} emails found.")

def save_results():
    if not found_emails:
        messagebox.showwarning("No Results", "No emails to export.")
        return
    folder = filedialog.askdirectory(initialdir=os.path.expanduser("~/Documents"))
    if not folder:
        return
    filename = f"Filtered_Emails_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    ws.append(["Subject", "Received", "Body", "Recipients", "CC", "Attachments"])
    for r in found_emails:
        ws.append([r["subject"], r["received"], r["body"], r["recipients"], r["cc"], r["attachments"]])
    wb.save(os.path.join(folder, filename))
    messagebox.showinfo("Export Complete", "Excel file saved successfully.")

# === Build GUI ===

root = Tk()
root.title("Outlook Email Search Tool")
root.configure(bg="#2e2e2e")

style = Style()
style.theme_use("clam")
style.configure(".", background="#2e2e2e", foreground="white", fieldbackground="#3e3e3e")
style.map("TCombobox", fieldbackground=[("readonly", "#3e3e3e")], foreground=[("readonly", "white")])

def label(row, text):
    Label(root, text=text, bg="#2e2e2e", fg="white").grid(row=row, column=0, sticky="w", pady=2)

def entry(row, width=50):
    e = Entry(root, width=width, bg="#3e3e3e", fg="white", insertbackground="white")
    e.grid(row=row, column=1, sticky="w", pady=2)
    return e

def dropdown(row, values, default):
    cb = Combobox(root, values=values, width=5, state="readonly")
    cb.set(default)
    cb.grid(row=row, column=2, sticky="w", padx=5)
    return cb

label(0, "Shared Mailbox (optional):")
mailbox_entry = entry(0)

label(1, "Subject contains:")
subject_include = entry(1)
subject_logic = dropdown(1, ["AND", "OR"], "AND")

label(2, "Subject NOT contains:")
subject_exclude = entry(2)
subject_exclude_logic = dropdown(2, ["AND", "OR"], "OR")

label(3, "Body contains:")
body_keywords = entry(3)
body_logic = dropdown(3, ["AND", "OR"], "OR")

label(4, "Attachment keywords:")
attachment_keywords = entry(4)
attachment_logic = dropdown(4, ["AND", "OR"], "OR")

label(5, "Sender keywords:")
sender_keywords = entry(5)
sender_logic = dropdown(5, ["AND", "OR"], "OR")

label(6, "Start Date:")
start_date = DateEntry(root, width=15, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
start_date.grid(row=6, column=1, sticky="w", pady=2)

label(7, "End Date:")
end_date = DateEntry(root, width=15, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
end_date.grid(row=7, column=1, sticky="w", pady=2)

label(8, "Max results (optional):")
limit = entry(8, width=10)

search_button = Button(root, text="Run Search", command=run_search, bg="#444", fg="white")
search_button.grid(row=9, column=0, pady=10)

Button(root, text="Export to Excel", command=save_results, bg="#444", fg="white").grid(row=9, column=1, pady=10)

status_label = Label(root, text="Ready", bg="#2e2e2e", fg="lightgray")
status_label.grid(row=10, column=0, columnspan=2, sticky="w")

label(11, "Results:")
output_box = ScrolledText(root, width=100, height=20, bg="#1e1e1e", fg="white", insertbackground="white")
output_box.grid(row=12, column=0, columnspan=3, pady=5)

found_emails = []
root.mainloop()

