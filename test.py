import pandas as pd

data = pd.read_excel("Emails.xlsx")

if 'Email' in data.columns:
    emails = list(data['Email'])
    email_collection = []
    for email in emails:
        if pd.isnull(email) == False:
            email_collection.append(email)
    emails = email_collection
else:
    print("doesn't exist")


