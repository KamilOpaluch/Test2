import win32com.client
import re

def expand_address_entry(address_entry):
    emails = []

    try:
        # Try to expand group
        members = address_entry.Members
        if members is not None:
            for i in range(1, members.Count + 1):
                member = members.Item(i)
                emails.extend(expand_address_entry(member))
        else:
            # Not a group: try to get Exchange email
            exch_user = address_entry.GetExchangeUser()
            if exch_user and exch_user.PrimarySmtpAddress:
                emails.append(exch_user.PrimarySmtpAddress)
            else:
                smtp_address = try_get_smtp_from_address_entry(address_entry)
                if smtp_address:
                    emails.append(smtp_address)
    except:
        smtp_address = try_get_smtp_from_address_entry(address_entry)
        if smtp_address:
            emails.append(smtp_address)

    return emails

def try_get_smtp_from_address_entry(address_entry):
    try:
        return address_entry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    except:
        return None

def resolve_and_expand(email_list):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    group_email_map = {}

    for item in email_list:
        try:
            recipient = namespace.CreateRecipient(item)
            if recipient.Resolve():
                address_entry = recipient.AddressEntry
                raw_emails = expand_address_entry(address_entry)

                # Clean each email
                clean_emails = list(filter(None, map(extract_email_only, raw_emails)))
                group_email_map[item] = sorted(set(clean_emails))
            else:
                group_email_map[item] = ["(Unresolved)"]
        except Exception as e:
            group_email_map[item] = [f"(Exception: {e})"]

    return group_email_map

def extract_email_only(item):
    match = re.search(r'<(.*?)>', item)
    if match:
        return match.group(1)
    elif '@' in item:
        return item.strip()
    return None

# Input: list of groups or addresses
input_emails = [
    "Finance Team",
    "hr@example.com"
]

# Expand and print
group_email_map = resolve_and_expand(input_emails)

print("\nGrouped Email List (copy-paste ready):\n")
for group, emails in group_email_map.items():
    print(f"{group}: {', '.join(emails)}")
