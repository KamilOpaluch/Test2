import win32com.client

def expand_address_entry(address_entry):
    result = []

    try:
        # Try to expand if it's a group
        members = address_entry.Members
        if members is not None:
            for i in range(1, members.Count + 1):
                member = members.Item(i)
                result.extend(expand_address_entry(member))
        else:
            # Not a group, get Exchange user or display name
            exch_user = address_entry.GetExchangeUser()
            if exch_user:
                result.append(f"{exch_user.Name} <{exch_user.PrimarySmtpAddress}>")
            else:
                result.append(address_entry.Name)
    except Exception as e:
        result.append(f"{address_entry.Name} (Error: {e})")

    return result

def resolve_and_expand(email_list):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    resolved_names = []

    for item in email_list:
        try:
            recipient = namespace.CreateRecipient(item)
            if recipient.Resolve():
                address_entry = recipient.AddressEntry
                resolved_names.extend(expand_address_entry(address_entry))
            else:
                resolved_names.append(f"{item} (Unresolved)")
        except Exception as e:
            resolved_names.append(f"{item} (Exception: {e})")

    return resolved_names

# Test with names or DLs you see in Outlook
input_emails = [
    "Finance Team",  # DL name from Global Address List
    "john.doe@example.com"
]

people = resolve_and_expand(input_emails)

print("Expanded members:")
for person in people:
    print(f" - {person}")
