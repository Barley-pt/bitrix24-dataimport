import requests
import openpyxl
import time

# Ask for file and webhook
excel_path = input("Enter path to Excel file: ").strip()
webhook = input("Enter Bitrix24 Webhook URL: ").strip()
delay = 1.5  # delay between requests

# Load workbook
wb = openpyxl.load_workbook(excel_path)
sheet = wb.active

# Read headers
headers = [cell.value for cell in sheet[1]]

# Column mapping
col = {h: i for i, h in enumerate(headers)}

# Loop through each row
for row in sheet.iter_rows(min_row=2, values_only=True):
    name = row[col['name']]
    phone = row[col['phone']]
    email = row[col['email']]
    company_name = row[col['company_name']]
    deal_title = row[col['deal_title']]
    deal_amount = row[col['deal_amount']]

    # 1. Create Contact
    contact_payload = {
        'fields': {
            'NAME': name,
            'PHONE': [{'VALUE': phone, 'VALUE_TYPE': 'WORK'}],
            'EMAIL': [{'VALUE': email, 'VALUE_TYPE': 'WORK'}]
        }
    }
    r = requests.post(f"{webhook}/crm.contact.add.json", json=contact_payload)
    contact_id = r.json().get('result')
    print(f"Created Contact ID: {contact_id}")
    time.sleep(delay)

    # 2. Create Company (optional)
    company_id = None
    if company_name:
        company_payload = {
            'fields': {
                'TITLE': company_name
            }
        }
        r = requests.post(f"{webhook}/crm.company.add.json", json=company_payload)
        company_id = r.json().get('result')
        print(f"Created Company ID: {company_id}")
        time.sleep(delay)

    # 3. Create Deal
    deal_payload = {
        'fields': {
            'TITLE': deal_title,
            'CONTACT_ID': contact_id,
            'COMPANY_ID': company_id,
            'OPPORTUNITY': deal_amount
        }
    }
    r = requests.post(f"{webhook}/crm.deal.add.json", json=deal_payload)
    deal_id = r.json().get('result')
    print(f"Created Deal ID: {deal_id}")
    time.sleep(delay)

print("✅ Import completed.")
