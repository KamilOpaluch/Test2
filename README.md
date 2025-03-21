import win32com.client

def expand_distribution_list(address_entry):
    result = []
    try:
        if address_entry.DisplayType == 1:  # Individual
            result.append(address_entry.Name)
        elif address_entry.DisplayType == 0:  # User (might be individual)
            exchUser = address_entry.GetExchangeUser()
            if exchUser:
                result.append(exchUser.Name)
        elif address_entry.DisplayType == 5:  # Exchange Distribution List
            members = address_entry.GetExchangeDistributionList().GetExchangeDistributionListMembers()
            if members:
                for member in members:
                    result.extend(expand_distribution_list(member))
            else:
                result.append(f"{address_entry.Name} (Empty group)")
        else:
            result.append(address_entry.Name)
    except Exception as e:
        result.append(f"{address_entry.Name} (Error: {e})")
    return result

def resolve_and_expand(email_list):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    resolved_names = []

    for email in email_list:
        address_entry = namespace.CreateRecipient(email).AddressEntry
        if address_entry is not None and address_entry.Resolved:
            resolved_names.extend(expand_distribution_list(address_entry))
        else:
            resolved_names.append(f"{email} (Unresolved)")
    
    return resolved_names

# Example list of email addresses and/or distribution list names
input_emails = [
    "john.doe@example.com",
    "Finance Team",  # example distribution group name
    "hr-department@example.com"
]

people = resolve_and_expand(input_emails)

print("Expanded members:")
for person in people:
    print(f" - {person}")
