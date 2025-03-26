import threading
import win32com.client
import pythoncom
import os
from datetime import datetime, timezone
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from tkcalendar import DateEntry

MAX_SCAN_LIMIT = 1000

def get_outlook_inbox(mailbox=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if mailbox:
        recipient = namespace.CreateRecipient(mailbox)
        recipient.Resolve()
        if recipient.Resolved:
            return namespace.GetSharedDefaultFolder(recipient, 6)
        else:
            raise Exception(f"Shared mailbox '{mailbox}' could not be resolved.")
    return namespace.GetDefaultFolder(6)

def match_keywords(text, keywords, logic):
    keywords = [k.strip() for k in keywords if k.strip()]
    if not keywords:
        return True
    text = text.lower()
    return all(k.lower() in text for k in keywords) if logic == "AND" else any(k.lower() in text for k in keywords)

def open_email_by_entryid(entryid):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mail_item = namespace.GetItemFromID(entryid)
    mail_item.Display()

def search_emails_thread(filters, on_result, on_done, on_debug):
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
        scanned = 0
        results = []

        for msg in messages:
            if scanned >= MAX_SCAN_LIMIT:
                on_debug(f"Stopped after scanning {MAX_SCAN_LIMIT} emails (max scan limit).")
                break
            scanned += 1

            try:
                subject = msg.Subject or ""
                body = msg.Body or ""
                sender = msg.SenderEmailAddress or ""
                attachments = [att.FileName for att in msg.Attachments]
                recipients = [msg.Recipients.Item(i + 1).Address for i in range(msg.Recipients.Count)]
                cc = msg.CC if msg.CC else ""
                received = msg.ReceivedTime

                if not match_keywords(subject, filters['subject_include'], filters['subject_logic']):
                    on_debug(f"Skipped: subject does not contain any of: {filters['subject_include']}")
                    continue
                if any(k.strip() for k in filters['subject_exclude']):
                    if match_keywords(subject, filters['subject_exclude'], filters['subject_exclude_logic']):
                        on_debug(f"Skipped: subject contains excluded keywords: {filters['subject_exclude']}")
                        continue
                if not match_keywords(body, filters['body_keywords'], filters['body_logic']):
                    on_debug("Skipped: body does not contain required keywords.")
                    continue
                if not match_keywords(sender, filters['sender_keywords'], filters['sender_logic']):
                    on_debug("Skipped: sender does not match provided filters.")
                    continue
                if any(k.strip() for k in filters['attachment_keywords']):
                    all_attachments = ' '.join(attachments).lower()
                    if not match_keywords(all_attachments, filters['attachment_keywords'], filters['attachment_logic']):
                        on_debug("Skipped: attachment filename does not match keywords.")
                        continue

                result = {
                    "subject": subject,
                    "received": received.strftime("%Y-%m-%d %H:%M"),
                    "body_preview": body[:50].replace("\n", " "),
                    "entryid": msg.EntryID
                }

                results.append(result)
                on_result(result)

                count += 1
                if filters['limit'] and count >= filters['limit']:
                    on_debug(f"Reached match limit: {filters['limit']}")
                    break

            except Exception as e:
                on_debug(f"Skipped due to error: {str(e)}")
                continue

        on_done(results)
    except Exception as e:
        on_done([], error=str(e))

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

        for row in tree.get_children():
            tree.delete(row)
        debug_box.delete("1.0", END)
        search_button.config(state=DISABLED)
        status_label.config(text="Searching...")

        global found_emails
        found_emails = []

        threading.Thread(
            target=search_emails_thread,
            args=(filters, on_result_found, on_search_done, on_debug_message),
            daemon=True
        ).start()

    except Exception as e:
        messagebox.showerror("Error", str(e))
        search_button.config(state=NORMAL)

def on_result_found(result):
    found_emails.append(result)
    tree.insert("", END, values=(result['received'], result['subject'], result['body_preview'], "Open"), tags=(result['entryid'],))

def on_search_done(results, error=None):
    search_button.config(state=NORMAL)
    if error:
        messagebox.showerror("Search Error", error)
        status_label.config(text="Search failed.")
    else:
        status_label.config(text=f"Search complete. {len(results)} emails found.")

def on_debug_message(msg):
    debug_box.insert(END, f"{msg}\n")
    debug_box.see(END)

def on_tree_click(event):
    item = tree.identify_row(event.y)
    if not item:
        return
    col = tree.identify_column(event.x)
    if col == '#4':  # 'Open' column
        entryid = tree.item(item, 'tags')[0]
        open_email_by_entryid(entryid)

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
    ws.append(["Subject", "Received", "Body Preview"])
    for r in found_emails:
        ws.append([r["subject"], r["received"], r["body_preview"]])
    wb.save(os.path.join(folder, filename))
    messagebox.showinfo("Export Complete", "Excel file saved successfully.")

# GUI
root = Tk()
root.title("Outlook Email Search Tool")
root.configure(bg="#2e2e2e")

style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", background="#2e2e2e", foreground="white")
style.configure("TButton", background="#3e3e3e", foreground="white", padding=5)
style.configure("Treeview", background="#1e1e1e", foreground="white", fieldbackground="#1e1e1e")
style.configure("Treeview.Heading", background="#333", foreground="white")
style.map("TButton", background=[("active", "#444")])

frm = Frame(root, bg="#2e2e2e")
frm.grid(row=0, column=0, sticky="w", padx=10, pady=10)

def add_label_entry_combo(r, label_txt):
    Label(frm, text=label_txt).grid(row=r, column=0, sticky="w")
    e = Entry(frm, width=50, bg="#3e3e3e", fg="white", insertbackground="white")
    e.grid(row=r, column=1, padx=5, pady=2)
    cb = Combobox(frm, values=["AND", "OR"], width=5, state="readonly")
    cb.set("OR")
    cb.grid(row=r, column=2, padx=5)
    return e, cb

Label(frm, text="Shared Mailbox (optional):").grid(row=0, column=0, sticky="w")
mailbox_entry = Entry(frm, width=50, bg="#3e3e3e", fg="white", insertbackground="white")
mailbox_entry.grid(row=0, column=1, columnspan=2, pady=2, sticky="w")

subject_include, subject_logic = add_label_entry_combo(1, "Subject contains:")
subject_exclude, subject_exclude_logic = add_label_entry_combo(2, "Subject NOT contains:")
body_keywords, body_logic = add_label_entry_combo(3, "Body contains:")
attachment_keywords, attachment_logic = add_label_entry_combo(4, "Attachment keywords:")
sender_keywords, sender_logic = add_label_entry_combo(5, "Sender keywords:")

Label(frm, text="Start Date:").grid(row=6, column=0, sticky="w")
start_date = DateEntry(frm, width=15, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
start_date.grid(row=6, column=1, sticky="w", pady=2)

Label(frm, text="End Date:").grid(row=7, column=0, sticky="w")
end_date = DateEntry(frm, width=15, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
end_date.grid(row=7, column=1, sticky="w", pady=2)

Label(frm, text="Max results:").grid(row=8, column=0, sticky="w")
limit = Entry(frm, width=10, bg="#3e3e3e", fg="white", insertbackground="white")
limit.grid(row=8, column=1, sticky="w", pady=2)

search_button = Button(frm, text="Run Search", command=run_search)
search_button.grid(row=9, column=0, pady=10)

Button(frm, text="Export to Excel", command=save_results).grid(row=9, column=1, pady=10)

status_label = Label(frm, text="Ready", fg="lightgray", bg="#2e2e2e")
status_label.grid(row=10, column=0, columnspan=2, sticky="w")

columns = ("received", "subject", "preview", "action")
tree = ttk.Treeview(root, columns=columns, show="headings", height=10)
tree.heading("received", text="Date")
tree.heading("subject", text="Subject")
tree.heading("preview", text="Body Preview")
tree.heading("action", text="Action")
tree.column("action", width=70, anchor="center")
tree.grid(row=11, column=0, sticky="nsew", padx=10, pady=5)
tree.bind("<ButtonRelease-1>", on_tree_click)

Label(root, text="Debug Log:", bg="#2e2e2e", fg="gray").grid(row=12, column=0, sticky="w", padx=10)
debug_box = Text(root, width=120, height=6, bg="#1e1e1e", fg="gray", insertbackground="white")
debug_box.grid(row=13, column=0, padx=10, pady=5)

found_emails = []
root.mainloop()
