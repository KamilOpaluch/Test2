import win32com.client

def expand_distribution_list(address_entry):
    result = []
    try:
        if address_entry.DisplayType == 1:  # Individual
            result.append(address_entry.Name)
        elif address_entry.DisplayType == 0:  # User
            exch_user = address_entry.GetExchangeUser()
            if exch_user:
                result.append(exch_user.Name)
        elif address_entry.DisplayType == 5:  # Distribution List
            dl = address_entry.GetExchangeDistributionList()
            if dl:
                members = dl.GetExchangeDistributionListMembers()
                if members:
                    for member in members:
                        result.extend(expand_distribution_list(member))
                else:
                    result.append(f"{address_entry.Name} (Empty group)")
            else:
                result.append(f"{address_entry.Name} (Not a DL)")
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
                resolved_names.extend(expand_distribution_list(address_entry))
            else:
                resolved_names.append(f"{item} (Unresolved)")
        except Exception as e:
            resolved_names.append(f"{item} (Exception: {e})")

    return resolved_names

# Example input
input_emails = [
    "john.doe@example.com",         # Individual
    "Finance Team",                 # Exchange DL name
    "hr-department@example.com"     # Email-based group
]

people = resolve_and_expand(input_emails)

print("Expanded members:")
for person in people:
    print(f" - {person}")
