import threading
import win32com.client
import pythoncom
import os
import re
import queue
from datetime import datetime, timezone
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from tkcalendar import DateEntry

MAX_SCAN_LIMIT = 1000
SEARCH_CANCEL_FLAG = False

def get_outlook_inbox(mailbox=None):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        if mailbox and mailbox.strip():
            try:
                recipient = namespace.CreateRecipient(mailbox)
                recipient.Resolve()
                if recipient.Resolved:
                    return namespace.GetSharedDefaultFolder(recipient, 6)
                else:
                    raise Exception(f"Shared mailbox '{mailbox}' could not be resolved.")
            except Exception as e:
                raise Exception(f"Error accessing shared mailbox '{mailbox}': {str(e)}")
        return namespace.GetDefaultFolder(6)  # Default inbox
    except Exception as e:
        raise Exception(f"Failed to access Outlook inbox: {str(e)}")

def match_keywords(text, keywords, logic):
    keywords = [k.strip() for k in keywords if k.strip()]
    if not keywords:
        return True
    if not text:
        return False
    text = text.lower()
    return all(k.lower() in text for k in keywords) if logic == "AND" else any(k.lower() in text for k in keywords)

def match_sender(sender_name, sender_email, keywords, logic):
    """Enhanced sender matching that checks both name and email"""
    keywords = [k.strip().lower() for k in keywords if k.strip()]
    if not keywords:
        return True
    
    # Prepare search fields
    search_fields = []
    if sender_name:
        search_fields.append(sender_name.lower())
        # Also try reversed name format (e.g., "Doe, John" vs "John Doe")
        if ', ' in sender_name:
            last, first = sender_name.split(', ', 1)
            search_fields.append(f"{first} {last}".lower())
    if sender_email:
        search_fields.append(sender_email.lower())
        # Also check email prefix (before @)
        if '@' in sender_email:
            search_fields.append(sender_email.split('@')[0].lower())
    
    # If no sender info available
    if not search_fields:
        return False
    
    # Perform matching
    if logic == "AND":
        return all(any(k in field for field in search_fields) for k in keywords)
    else:  # OR logic
        return any(any(k in field for field in search_fields) for k in keywords)

def open_email_by_entryid(entryid):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(entryid)
        mail_item.Display()
    except Exception as e:
        messagebox.showerror("Error", f"Could not open email: {str(e)}")

def clean_text(text):
    """Remove extra spaces and newlines from text"""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text.strip())
    return text

def search_emails_thread(filters, on_result, on_done, on_debug, cancel_queue):
    global SEARCH_CANCEL_FLAG
    SEARCH_CANCEL_FLAG = False
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
            # Check for cancel request
            try:
                if not cancel_queue.empty():
                    cancel_queue.get()
                    SEARCH_CANCEL_FLAG = True
                    break
            except:
                pass
            
            if SEARCH_CANCEL_FLAG or scanned >= MAX_SCAN_LIMIT:
                on_debug(f"Stopped after scanning {scanned} emails.")
                break
            scanned += 1

            try:
                # Safely get email properties with defaults
                subject = getattr(msg, 'Subject', '') or ""
                body = getattr(msg, 'Body', '') or ""
                sender_email = getattr(msg, 'SenderEmailAddress', '') or ""
                sender_name = getattr(msg, 'SenderName', '') or ""
                
                # Handle recipients
                recipients = []
                if hasattr(msg, 'Recipients') and msg.Recipients:
                    for i in range(msg.Recipients.Count):
                        try:
                            recipients.append(msg.Recipients.Item(i + 1).Address)
                        except:
                            continue
                recipients_str = "; ".join(recipients) if recipients else ""
                
                # Handle CC
                cc = getattr(msg, 'CC', '') or ""
                
                # Handle attachments
                attachments = []
                if hasattr(msg, 'Attachments') and msg.Attachments:
                    for att in msg.Attachments:
                        try:
                            attachments.append(att.FileName)
                        except:
                            continue
                
                received = getattr(msg, 'ReceivedTime', datetime.now())
                importance = getattr(msg, 'Importance', 1)  # Default to normal importance
                read_status = not getattr(msg, 'UnRead', True)

                # Apply filters
                if not match_keywords(subject, filters['subject_include'], filters['subject_logic']):
                    continue
                if any(k.strip() for k in filters['subject_exclude']):
                    if match_keywords(subject, filters['subject_exclude'], filters['subject_exclude_logic']):
                        continue
                if not match_keywords(body, filters['body_keywords'], filters['body_logic']):
                    continue
                if not match_sender(sender_name, sender_email, filters['sender_keywords'], filters['sender_logic']):
                    continue
                if any(k.strip() for k in filters['attachment_keywords']):
                    all_attachments = ' '.join(attachments).lower()
                    if not match_keywords(all_attachments, filters['attachment_keywords'], filters['attachment_logic']):
                        continue

                result = {
                    "subject": subject,
                    "received": received.strftime("%Y-%m-%d %H:%M"),
                    "body_preview": clean_text(body)[:100] + ("..." if len(body) > 100 else ""),
                    "body_full": body,
                    "entryid": msg.EntryID,
                    "from": sender_name,
                    "to": recipients_str,
                    "cc": cc,
                    "attachments": "; ".join(attachments) if attachments else "",
                    "importance": importance,
                    "read": read_status
                }

                results.append(result)
                on_result(result)

                count += 1
                if filters['limit'] and count >= filters['limit']:
                    on_debug(f"Reached match limit: {filters['limit']}")
                    break

            except Exception as e:
                on_debug(f"Skipped email due to error: {str(e)}")
                continue

        on_done(results, cancelled=SEARCH_CANCEL_FLAG)
    except Exception as e:
        on_done([], error=f"Search failed: {str(e)}")

def run_search():
    global SEARCH_CANCEL_FLAG, cancel_queue
    SEARCH_CANCEL_FLAG = False
    cancel_queue = queue.Queue()
    
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
        cancel_button.config(state=NORMAL)
        status_label.config(text="Searching...")
        progress_bar.start()

        global found_emails
        found_emails = []

        threading.Thread(
            target=search_emails_thread,
            args=(filters, on_result_found, on_search_done, on_debug_message, cancel_queue),
            daemon=True
        ).start()

    except Exception as e:
        messagebox.showerror("Error", str(e))
        search_button.config(state=NORMAL)
        cancel_button.config(state=DISABLED)

def cancel_search():
    global SEARCH_CANCEL_FLAG
    SEARCH_CANCEL_FLAG = True
    try:
        cancel_queue.put(True)
    except:
        pass
    cancel_button.config(state=DISABLED)

def on_result_found(result):
    found_emails.append(result)
    importance_icon = "↑" if result['importance'] == 2 else ("↓" if result['importance'] == 0 else "•")
    read_icon = "✓" if result['read'] else "✕"
    tree.insert("", END, values=(
        result['received'], 
        result['subject'], 
        result['from'],
        result['body_preview'], 
        importance_icon,
        read_icon,
        "Open"
    ), tags=(result['entryid'],))

def on_search_done(results, error=None, cancelled=False):
    search_button.config(state=NORMAL)
    cancel_button.config(state=DISABLED)
    progress_bar.stop()
    
    if error:
        messagebox.showerror("Search Error", error)
        status_label.config(text="Search failed.")
    elif cancelled:
        status_label.config(text=f"Search cancelled. {len(results)} emails found before cancellation.")
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
    if col == '#7':  # 'Open' column
        entryid = tree.item(item, 'tags')[0]
        open_email_by_entryid(entryid)

def save_results():
    if not found_emails:
        messagebox.showwarning("No Results", "No emails to export.")
        return
    
    file = filedialog.asksaveasfilename(
        initialdir=os.path.expanduser("~/Documents"),
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=f"Filtered_Emails_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )
    
    if not file:
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Emails"
        
        headers = [
            "Subject", "Received", "From", "To", "CC", 
            "Body", "Attachments", "Importance", "Read Status"
        ]
        ws.append(headers)
        
        for r in found_emails:
            ws.append([
                r["subject"],
                r["received"],
                r["from"],
                r["to"],
                r["cc"],
                r["body_full"],
                r["attachments"],
                "High" if r["importance"] == 2 else ("Low" if r["importance"] == 0 else "Normal"),
                "Read" if r["read"] else "Unread"
            ])
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column[0].column_letter].width = adjusted_width
        
        wb.save(file)
        messagebox.showinfo("Export Complete", f"Excel file saved successfully:\n{file}")
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to save file:\n{str(e)}")

# GUI Setup
root = Tk()
root.title("Outlook Email Search Tool")
root.geometry("1100x800")
root.minsize(900, 600)
root.configure(bg="#f0f0f0")

# Modern styling
style = ttk.Style()
style.theme_use("clam")

# Configure styles
style.configure(".", background="#f0f0f0", foreground="#333333", font=('Segoe UI', 9))
style.configure("TFrame", background="#f0f0f0")
style.configure("TLabel", background="#f0f0f0", foreground="#333333")
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

style.configure("Vertical.TScrollbar", background="#e0e0e0")

# Main frame
main_frame = ttk.Frame(root, padding="15")
main_frame.grid(row=0, column=0, sticky=(N, S, E, W))

# Configure grid weights
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

# Search parameters frame
params_frame = ttk.LabelFrame(main_frame, text="Search Parameters", padding=(10, 5))
params_frame.grid(row=0, column=0, columnspan=2, sticky=(N, S, E, W), pady=(0, 10))
params_frame.columnconfigure(1, weight=1)

# Form fields
ttk.Label(params_frame, text="Shared Mailbox (optional):").grid(row=0, column=0, sticky=W, pady=3)
mailbox_entry = ttk.Entry(params_frame, width=50)
mailbox_entry.grid(row=0, column=1, sticky=W, pady=3, columnspan=2)

def create_filter_row(row, label_text):
    ttk.Label(params_frame, text=label_text).grid(row=row, column=0, sticky=W, pady=3)
    entry = ttk.Entry(params_frame, width=50)
    entry.grid(row=row, column=1, sticky=W, pady=3)
    logic = ttk.Combobox(params_frame, values=["AND", "OR"], width=5, state="readonly")
    logic.set("OR")
    logic.grid(row=row, column=2, padx=5, sticky=W)
    return entry, logic

subject_include, subject_logic = create_filter_row(1, "Subject contains:")
subject_exclude, subject_exclude_logic = create_filter_row(2, "Subject NOT contains:")
body_keywords, body_logic = create_filter_row(3, "Body contains:")
attachment_keywords, attachment_logic = create_filter_row(4, "Attachment keywords:")
sender_keywords, sender_logic = create_filter_row(5, "Sender keywords:")

# Date range frame
date_frame = ttk.Frame(params_frame)
date_frame.grid(row=6, column=0, columnspan=3, sticky=W, pady=5)

ttk.Label(date_frame, text="Start Date:").grid(row=0, column=0, sticky=W, padx=(0, 10))
start_date = DateEntry(date_frame, width=15, background="#4CAF50", foreground="white", 
                      borderwidth=1, date_pattern="yyyy-mm-dd")
start_date.grid(row=0, column=1, sticky=W, padx=(0, 20))

ttk.Label(date_frame, text="End Date:").grid(row=0, column=2, sticky=W, padx=(0, 10))
end_date = DateEntry(date_frame, width=15, background="#4CAF50", foreground="white", 
                    borderwidth=1, date_pattern="yyyy-mm-dd")
end_date.grid(row=0, column=3, sticky=W)

# Limit and buttons frame
limit_frame = ttk.Frame(params_frame)
limit_frame.grid(row=7, column=0, columnspan=3, sticky=W, pady=5)

ttk.Label(limit_frame, text="Max results:").grid(row=0, column=0, sticky=W, padx=(0, 10))
limit = ttk.Entry(limit_frame, width=10)
limit.grid(row=0, column=1, sticky=W)

# Buttons
button_frame = ttk.Frame(main_frame)
button_frame.grid(row=1, column=0, columnspan=2, pady=(0, 10), sticky=W)

search_button = ttk.Button(button_frame, text="Run Search", command=run_search)
search_button.grid(row=0, column=0, padx=5)

cancel_button = ttk.Button(button_frame, text="Cancel", command=cancel_search, state=DISABLED)
cancel_button.grid(row=0, column=1, padx=5)

export_button = ttk.Button(button_frame, text="Export to Excel", command=save_results)
export_button.grid(row=0, column=2, padx=5)

# Progress bar
progress_bar = ttk.Progressbar(button_frame, mode='indeterminate', length=200)
progress_bar.grid(row=0, column=3, padx=10)

# Status label
status_label = ttk.Label(button_frame, text="Ready", foreground="#666666")
status_label.grid(row=0, column=4, padx=10)

# Results frame
results_frame = ttk.LabelFrame(main_frame, text="Search Results", padding=(10, 5))
results_frame.grid(row=2, column=0, columnspan=2, sticky=(N, S, E, W), pady=(0, 10))
results_frame.columnconfigure(0, weight=1)
results_frame.rowconfigure(0, weight=1)

# Results table
columns = ("received", "subject", "from", "preview", "importance", "read", "action")
tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=15)
tree.heading("received", text="Date")
tree.heading("subject", text="Subject")
tree.heading("from", text="From")
tree.heading("preview", text="Body Preview")
tree.heading("importance", text="!")
tree.heading("read", text="✓")
tree.heading("action", text="Action")

# Configure column widths
tree.column("received", width=150, anchor=W)
tree.column("subject", width=200, anchor=W)
tree.column("from", width=150, anchor=W)
tree.column("preview", width=300, anchor=W)
tree.column("importance", width=30, anchor=CENTER)
tree.column("read", width=30, anchor=CENTER)
tree.column("action", width=70, anchor=CENTER)

tree.grid(row=0, column=0, sticky=(N, S, E, W))
tree.bind("<ButtonRelease-1>", on_tree_click)

# Add scrollbars
v_scroll = ttk.Scrollbar(results_frame, orient=VERTICAL, command=tree.yview)
tree.configure(yscrollcommand=v_scroll.set)
v_scroll.grid(row=0, column=1, sticky=(N, S))

h_scroll = ttk.Scrollbar(results_frame, orient=HORIZONTAL, command=tree.xview)
tree.configure(xscrollcommand=h_scroll.set)
h_scroll.grid(row=1, column=0, sticky=(E, W))

# Debug log frame
debug_frame = ttk.LabelFrame(main_frame, text="Debug Log", padding=(10, 5))
debug_frame.grid(row=3, column=0, columnspan=2, sticky=(N, S, E, W))
debug_frame.columnconfigure(0, weight=1)
debug_frame.rowconfigure(0, weight=1)

debug_box = Text(debug_frame, width=120, height=6, bg="white", fg="#333333", 
                insertbackground="#333333", wrap=WORD, padx=5, pady=5,
                highlightbackground="#cccccc", highlightcolor="#cccccc", highlightthickness=1)
debug_box.grid(row=0, column=0, sticky=(N, S, E, W))

debug_scroll = ttk.Scrollbar(debug_frame, orient=VERTICAL, command=debug_box.yview)
debug_box.configure(yscrollcommand=debug_scroll.set)
debug_scroll.grid(row=0, column=1, sticky=(N, S))

# Configure weights for resizing
main_frame.rowconfigure(2, weight=1)
main_frame.rowconfigure(3, weight=0)
main_frame.columnconfigure(0, weight=1)

found_emails = []
cancel_queue = None
root.mainloop()
