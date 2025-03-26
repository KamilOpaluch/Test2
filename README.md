def search_matching_emails(inbox_folder, keyword, after_date, target_email):
    messages = inbox_folder.Items

    # Apply server-side restriction to speed up
    restriction = "[SentOn] >= '" + after_date.strftime("%m/%d/%Y %H:%M %p") + "'"
    restricted_items = messages.Restrict(restriction)
    restricted_items.Sort("[SentOn]", True)

    results = []

    for msg in restricted_items:
        try:
            subject = msg.Subject
            if keyword.lower() not in subject.lower():
                continue

            # Filter by recipients (contains the shared mailbox)
            recipients = [msg.Recipients.Item(i + 1).AddressEntry.Address.lower()
                          for i in range(msg.Recipients.Count)]
            if not any(target_email.lower() in r for r in recipients):
                continue

            results.append((subject, msg.SentOn.strftime("%Y-%m-%d %H:%M")))
        except Exception:
            continue

    return results
