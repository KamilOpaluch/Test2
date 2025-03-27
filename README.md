def search_emails_thread(filters, on_result, on_done, on_debug, cancel_queue):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    try:
        inbox = get_outlook_inbox(filters.get('mailbox'))
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        
        # Build efficient restriction
        restriction_parts = [
            "[ReceivedTime] >= '{}'".format(filters['start_date'].strftime("%m/%d/%Y %H:%M")),
            "[ReceivedTime] <= '{}'".format(filters['end_date'].strftime("%m/%d/%Y %H:%M"))
        ]
        
        if filters['subject_include']:
            subj_include = [k for k in filters['subject_include'] if k.strip()]
            if subj_include:
                subj_restriction = " OR ".join(
                    f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{k}%'" 
                    for k in subj_include
                )
                restriction_parts.append(f"({subj_restriction})")
        
        final_restriction = " AND ".join(restriction_parts)
        messages = items.Restrict(final_restriction)
        
        # Pre-compile patterns for faster matching
        sender_patterns = [re.compile(re.escape(k.lower())) 
                         for k in filters['sender_keywords'] if k.strip()]
        
        count = 0
        batch = []
        BATCH_SIZE = 50
        
        for i, msg in enumerate(messages):
            if SEARCH_CANCEL_FLAG or i >= MAX_SCAN_LIMIT:
                break
                
            try:
                # Early filtering
                subject = msg.Subject or ""
                if not any(k.lower() in subject.lower() 
                          for k in filters['subject_include'] if k.strip()):
                    continue
                    
                # Process in batches
                batch.append(msg)
                if len(batch) >= BATCH_SIZE:
                    process_batch(batch, filters, on_result, on_debug)
                    batch = []
                    
                count += 1
                if filters['limit'] and count >= filters['limit']:
                    break
                    
            except Exception as e:
                on_debug(f"Skipped error: {str(e)}")
                continue
                
        # Process remaining messages
        if batch:
            process_batch(batch, filters, on_result, on_debug)
            
        on_done(count, cancelled=SEARCH_CANCEL_FLAG)
        
    except Exception as e:
        on_done(0, error=str(e))
    finally:
        pythoncom.CoUninitialize()

def process_batch(messages, filters, on_result, on_debug):
    results = []
    for msg in messages:
        try:
            subject = msg.Subject or ""
            body = msg.Body or ""
            sender = msg.SenderEmailAddress or ""
            
            # Additional filtering
            if not match_keywords(subject, filters['subject_include'], filters['subject_logic']):
                continue
                
            result = {
                "subject": subject,
                "received": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
                "body": body,
                "from": msg.SenderName or ""
            }
            results.append(result)
            on_result(result)
            
        except Exception as e:
            on_debug(f"Batch error: {str(e)}")
            continue
