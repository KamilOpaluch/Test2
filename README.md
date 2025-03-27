def search_emails_thread(filters, on_result, on_done, on_debug, cancel_queue):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    try:
        inbox = get_outlook_inbox(filters.get('mailbox'))
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        
        # Build proper restriction query
        restriction_parts = [
            "[ReceivedTime] >= '{}'".format(filters['start_date'].strftime("%m/%d/%Y %H:%M %p")),
            "[ReceivedTime] <= '{}'".format(filters['end_date'].strftime("%m/%d/%Y %H:%M %p"))
        ]
        
        # Add subject filters only if they exist
        if filters['subject_include'] and any(k.strip() for k in filters['subject_include']):
            subj_conditions = []
            for k in filters['subject_include']:
                if k.strip():
                    # Properly escape single quotes in subject
                    clean_k = k.replace("'", "''")
                    subj_conditions.append("""@SQL="urn:schemas:httpmail:subject" LIKE '%{}%'""".format(clean_k))
            
            if subj_conditions:
                subj_restriction = " OR ".join(subj_conditions)
                if filters['subject_logic'] == "AND":
                    subj_restriction = "(" + " AND ".join(subj_conditions) + ")"
                restriction_parts.append(subj_restriction)
        
        # Build final restriction
        final_restriction = " AND ".join(restriction_parts)
        on_debug(f"Using restriction: {final_restriction}")
        
        try:
            messages = items.Restrict(final_restriction)
        except Exception as e:
            on_debug(f"Restrict failed, falling back to full scan: {str(e)}")
            messages = items  # Fallback to all items if restriction fails
        
        count = 0
        scanned = 0
        batch_size = 50
        batch = []
        
        for msg in messages:
            if SEARCH_CANCEL_FLAG or scanned >= MAX_SCAN_LIMIT:
                break
                
            scanned += 1
            batch.append(msg)
            
            if len(batch) >= batch_size:
                count += process_batch(batch, filters, on_result, on_debug)
                batch = []
                if filters['limit'] and count >= filters['limit']:
                    break
        
        # Process remaining messages in batch
        if batch and (not filters['limit'] or count < filters['limit']):
            count += process_batch(batch, filters, on_result, on_debug)
            
        on_done(count, cancelled=SEARCH_CANCEL_FLAG)
        
    except Exception as e:
        on_done(0, error=str(e))
    finally:
        pythoncom.CoUninitialize()

def process_batch(messages, filters, on_result, on_debug):
    count = 0
    for msg in messages:
        try:
            subject = msg.Subject or ""
            body = msg.Body or ""
            sender = msg.SenderEmailAddress or ""
            
            # Skip if doesn't match subject includes
            if not match_keywords(subject, filters['subject_include'], filters['subject_logic']):
                continue
                
            # Skip if matches subject excludes
            if any(k.strip() for k in filters['subject_exclude']):
                if match_keywords(subject, filters['subject_exclude'], filters['subject_exclude_logic']):
                    continue
            
            # Other filters (body, sender, attachments)...
            
            result = {
                "subject": subject,
                "received": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
                "body": body,
                "from": msg.SenderName or "",
                "entryid": msg.EntryID
            }
            on_result(result)
            count += 1
            
        except Exception as e:
            on_debug(f"Error processing message: {str(e)}")
            
    return count
