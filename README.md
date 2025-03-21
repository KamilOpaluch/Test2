import re

# Filter to extract only valid emails (inside <> or plain)
def extract_email_only(item):
    match = re.search(r'<(.*?)>', item)
    if match:
        return match.group(1)
    elif '@' in item:
        return item.strip()
    return None

clean_emails = list(filter(None, map(extract_email_only, emails)))
unique_emails = sorted(set(clean_emails))

print("\nCopy-paste ready email list:")
print(", ".join(unique_emails))
