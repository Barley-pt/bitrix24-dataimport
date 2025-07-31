# Bitrix24 Importer

A powerful, user-friendly Python tool to import Deals (with linked Contacts) into Bitrix24 from Excel or CSV files.  
Supports field mapping, flexible deduplication, and works with custom Bitrix24 fields.

---

## Features

- Map Excel/CSV columns to Bitrix24 Deal and Contact fields (including custom fields)
- Bulk import Deals and Contacts, automatically linking them
- Deduplication: Avoid duplicate contacts by matching on Email, Phone, Name, or any mapped field
- Scrollable, resizable mapping window for easy large imports
- Verbose progress and clear error messages
- Supports both `.xlsx`/`.xls` and `.csv` formats
- No data is imported until you approve field mappings

---

## How it Works

1. **Webhook Setup:**  
   Enter your Bitrix24 REST webhook (needs CRM permissions).

2. **Fetch Fields:**  
   The tool fetches all available fields for Contacts and Deals, including user fields.

3. **Pipeline Selection:**  
   Choose the Bitrix24 Deal pipeline (category) for the import.

4. **File Selection:**  
   Choose your Excel or CSV file with your data.

5. **Field Mapping:**  
   Map your file’s columns to Bitrix24 Contact and Deal fields in easy, scrollable UIs.

6. **Deduplication Selection:**  
   Select the column to deduplicate contacts (e.g., Email, Phone, Name).  
   If left blank, will use the first "email"-like column or the first mapped field as fallback.

7. **Deal Contact Link:**  
   Select which Deal field will receive the Contact ID (usually `CONTACT_ID`).

8. **Import:**  
   For each row:  
   - Contact is created or updated (using deduplication)  
   - Deal is created, linked to the correct contact

9. **Done:**  
   You'll see a summary dialog when the import is finished.

---

## Requirements

- Python 3.8 or newer
- Packages:  
  `requests`, `pandas`, `openpyxl`  
  *(install via `pip install requests pandas openpyxl`)*

---

## Usage

```bash
python bitrix24\ importer.py

