import threading
import win32com.client
import pythoncom
import os
import re
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

def clean_text(text):
    """Remove extra spaces and newlines from text"""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text.strip())
    return text[:100] + "..." if len(text) > 100 else text

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
                sender_name = msg.SenderName or ""
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
                    "body_preview": clean_text(body),
                    "entryid": msg.EntryID,
                    "from": sender_name
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
    tree.insert("", END, values=(
        result['received'], 
        result['subject'], 
        result['from'],
        result['body_preview'], 
        "Open"
    ), tags=(result['entryid'],))

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
    if col == '#5':  # 'Open' column
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
    ws.append(["Subject", "Received", "From", "Body Preview"])
    for r in found_emails:
        ws.append([r["subject"], r["received"], r["from"], r["body_preview"]])
    wb.save(os.path.join(folder, filename))
    messagebox.showinfo("Export Complete", "Excel file saved successfully.")

# GUI Setup
root = Tk()
root.title("Outlook Email Search Tool")
root.geometry("1000x700")
root.configure(bg="#f0f0f0")

# Modern styling
style = ttk.Style()
style.theme_use("clam")

# Configure styles
style.configure("TFrame", background="#f0f0f0")
style.configure("TLabel", background="#f0f0f0", foreground="#333333", font=('Segoe UI', 9))
style.configure("TButton", 
                background="#4CAF50", 
                foreground="white", 
                font=('Segoe UI', 9, 'bold'),
                borderwidth=1,
                focusthickness=3,
                focuscolor='none')
style.map("TButton",
          background=[('active', '#45a049'), ('pressed', '#3d8b40')],
          foreground=[('pressed', 'white'), ('active', 'white')])

style.configure("TEntry", 
                fieldbackground="white", 
                foreground="#333333",
                insertcolor="#333333",
                bordercolor="#cccccc",
                lightcolor="#cccccc",
                darkcolor="#cccccc")

style.configure("Treeview", 
                background="white", 
                foreground="#333333", 
                fieldbackground="white",
                rowheight=25,
                font=('Segoe UI', 9))
style.configure("Treeview.Heading", 
                background="#4CAF50", 
                foreground="white",
                font=('Segoe UI', 9, 'bold'))
style.map("Treeview", 
          background=[('selected', '#4CAF50')],
          foreground=[('selected', 'white')])

style.configure("TCombobox", 
                fieldbackground="white", 
                foreground="#333333",
                selectbackground="#e6e6e6")

# Main frame
main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky=(N, S, E, W))

# Configure grid weights
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

# Form fields
ttk.Label(main_frame, text="Shared Mailbox (optional):").grid(row=0, column=0, sticky=W, pady=2)
mailbox_entry = ttk.Entry(main_frame, width=50)
mailbox_entry.grid(row=0, column=1, sticky=W, pady=2)

def create_filter_row(row, label_text):
    ttk.Label(main_frame, text=label_text).grid(row=row, column=0, sticky=W, pady=2)
    entry = ttk.Entry(main_frame, width=50)
    entry.grid(row=row, column=1, sticky=W, pady=2)
    logic = ttk.Combobox(main_frame, values=["AND", "OR"], width=5, state="readonly")
    logic.set("OR")
    logic.grid(row=row, column=2, padx=5, sticky=W)
    return entry, logic

subject_include, subject_logic = create_filter_row(1, "Subject contains:")
subject_exclude, subject_exclude_logic = create_filter_row(2, "Subject NOT contains:")
body_keywords, body_logic = create_filter_row(3, "Body contains:")
attachment_keywords, attachment_logic = create_filter_row(4, "Attachment keywords:")
sender_keywords, sender_logic = create_filter_row(5, "Sender keywords:")

# Date range
ttk.Label(main_frame, text="Start Date:").grid(row=6, column=0, sticky=W, pady=2)
start_date = DateEntry(main_frame, width=15, background="#4CAF50", foreground="white", 
                      borderwidth=1, date_pattern="yyyy-mm-dd")
start_date.grid(row=6, column=1, sticky=W, pady=2)

ttk.Label(main_frame, text="End Date:").grid(row=7, column=0, sticky=W, pady=2)
end_date = DateEntry(main_frame, width=15, background="#4CAF50", foreground="white", 
                    borderwidth=1, date_pattern="yyyy-mm-dd")
end_date.grid(row=7, column=1, sticky=W, pady=2)

# Max results
ttk.Label(main_frame, text="Max results:").grid(row=8, column=0, sticky=W, pady=2)
limit = ttk.Entry(main_frame, width=10)
limit.grid(row=8, column=1, sticky=W, pady=2)

# Buttons
button_frame = ttk.Frame(main_frame)
button_frame.grid(row=9, column=0, columnspan=3, pady=10)
search_button = ttk.Button(button_frame, text="Run Search", command=run_search)
search_button.grid(row=0, column=0, padx=5)
export_button = ttk.Button(button_frame, text="Export to Excel", command=save_results)
export_button.grid(row=0, column=1, padx=5)

# Status label
status_label = ttk.Label(main_frame, text="Ready", foreground="#666666")
status_label.grid(row=10, column=0, columnspan=3, sticky=W)

# Results table
columns = ("received", "subject", "from", "preview", "action")
tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=10)
tree.heading("received", text="Date")
tree.heading("subject", text="Subject")
tree.heading("from", text="From")
tree.heading("preview", text="Body Preview")
tree.heading("action", text="Action")

# Configure column widths
tree.column("received", width=150, anchor=W)
tree.column("subject", width=200, anchor=W)
tree.column("from", width=150, anchor=W)
tree.column("preview", width=300, anchor=W)
tree.column("action", width=70, anchor=CENTER)

tree.grid(row=11, column=0, columnspan=3, sticky=(N, S, E, W), pady=5)
tree.bind("<ButtonRelease-1>", on_tree_click)

# Add scrollbar
scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.grid(row=11, column=3, sticky=(N, S))

# Debug log
ttk.Label(main_frame, text="Debug Log:", foreground="#666666").grid(row=12, column=0, sticky=W, pady=(10, 0))
debug_box = Text(main_frame, width=120, height=6, bg="white", fg="#333333", 
                insertbackground="#333333", wrap=WORD, padx=5, pady=5,
                highlightbackground="#cccccc", highlightcolor="#cccccc", highlightthickness=1)
debug_box.grid(row=13, column=0, columnspan=3, sticky=(N, S, E, W), pady=(0, 10))

# Configure weights for resizing
main_frame.rowconfigure(11, weight=1)
main_frame.rowconfigure(13, weight=1)
main_frame.columnconfigure(1, weight=1)

found_emails = []
root.mainloop()
