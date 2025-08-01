import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import requests
import pandas as pd
import os
import sys
import csv
log_filename = "bitrix24_import_log.csv"
logfile = open(log_filename, mode="w", newline='', encoding="utf-8")
logwriter = csv.writer(logfile)
logwriter.writerow([
    "Row", "Contact_Dedup_Value", "Contact_Payload", "Contact_ID", "Contact_Result",
    "Deal_Payload", "Deal_ID", "Deal_Result"
])

# --- Helper for pretty Bitrix24 field labels ---
def field_label(fid, fdata):
    title = (
        fdata.get("listLabel")
        or fdata.get("formLabel")
        or fdata.get("filterLabel")
        or fdata.get("title")
        or fid
    )
    label = f"{title} ({fid})" if fid.startswith("UF_CRM") else title
    if fdata.get('type') == 'enumeration' and 'items' in fdata:
        try:
            if isinstance(fdata['items'], dict):
                enum_labels = list(fdata['items'].values())
            elif isinstance(fdata['items'], list):
                enum_labels = [item['VALUE'] for item in fdata['items']]
            else:
                enum_labels = []
            if enum_labels:
                label += " - [" + ", ".join(enum_labels) + "]"
        except Exception:
            pass
    return label

# --- GUI Mapping Window with Scroll ---
def mapping_window(columns, b24_fields, title):
    mapping = {}
    window = tk.Toplevel()
    window.title(title)
    window.geometry("700x500")
    window.minsize(400, 300)
    window.resizable(True, True)
    canvas = tk.Canvas(window)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    vscrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=canvas.yview)
    vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.configure(yscrollcommand=vscrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")
    dropdown_vars = []
    b24_choices = [""] + [field_label(fid, fdata) for fid, fdata in sorted(b24_fields.items(), key=lambda x: field_label(x[0], x[1]).lower())]
    b24_id_map = {field_label(fid, fdata): fid for fid, fdata in b24_fields.items()}
    tk.Label(frame, text="Excel Column", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tk.Label(frame, text="Bitrix24 Field", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=5)
    for i, col in enumerate(columns):
        tk.Label(frame, text=col, anchor="w").grid(row=i+1, column=0, sticky="w", padx=5, pady=3)
        var = tk.StringVar()
        dropdown = ttk.Combobox(frame, textvariable=var, values=b24_choices, state="readonly", width=60)
        dropdown.grid(row=i+1, column=1, sticky="w", padx=5, pady=3)
        dropdown_vars.append(var)
    def submit():
        for i, col in enumerate(columns):
            selected = dropdown_vars[i].get()
            fid = b24_id_map.get(selected)
            if fid:
                mapping[col] = fid
        window.destroy()
    submit_btn = tk.Button(frame, text="Submit", command=submit)
    submit_btn.grid(row=len(columns) + 2, column=0, columnspan=2, pady=20)
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
    canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))
    window.grab_set()
    window.wait_window()
    return mapping


# --- Fetch fields/pipelines from Bitrix24 ---
def fetch_fields(webhook, entity):
    resp = requests.get(f"{webhook}{entity}.fields.json")
    resp.raise_for_status()
    return resp.json()["result"]

def fetch_pipelines(webhook):
    resp = requests.get(f"{webhook}crm.dealcategory.list.json")
    resp.raise_for_status()
    return resp.json()["result"]

def sanitize_payload(data):
    """Recursively convert pandas Timestamps and other non-serializable objects to strings."""
    import pandas as pd
    import datetime
    sanitized = {}
    for k, v in data.items():
        if isinstance(v, pd.Timestamp):
            sanitized[k] = v.strftime("%Y-%m-%d")  # Or use .isoformat()
        elif isinstance(v, (datetime.date, datetime.datetime)):
            sanitized[k] = v.isoformat()
        elif pd.isnull(v):
            continue  # skip NaN/None
        else:
            sanitized[k] = v
    return sanitized

# --- Deduplication logic: find existing contact ---
def find_existing_contact(webhook, dedupe_field, value):
    if not value:
        return None
    filter_key = dedupe_field
    # For standard fields, Bitrix wants EMAIL/PHONE, for custom use the field code
    if dedupe_field.upper() == "EMAIL":
        filter_dict = {"filter[EMAIL]": value}
    elif dedupe_field.upper() == "PHONE":
        filter_dict = {"filter[PHONE]": value}
    elif dedupe_field.upper() == "NAME":
        filter_dict = {"filter[NAME]": value}
    else:
        filter_dict = {f"filter[{dedupe_field}]": value}
    r = requests.get(f"{webhook}crm.contact.list.json", params=filter_dict)
    data = r.json()
    if data.get('result') and len(data['result']) > 0:
        return data['result'][0]['ID']
    return None

# --- Create Bitrix24 Contact ---
def create_contact(webhook, contact_payload):
    r = requests.post(f"{webhook}crm.contact.add.json", json={"fields": contact_payload, "params": {"REGISTER_SONET_EVENT": "N"}})
    data = r.json()
    if 'result' in data:
        return data['result']
    else:
        print(f"Failed to create contact: {data}")
        return None

# --- Create Bitrix24 Deal ---
def create_deal(webhook, deal_payload):
    r = requests.post(f"{webhook}crm.deal.add.json", json={"fields": deal_payload, "params": {"REGISTER_SONET_EVENT": "N"}})
    data = r.json()
    if 'result' in data:
        return data['result']
    else:
        print(f"Failed to create deal: {data}")
        return None

# --- Main program ---
def main():
    # 1. Get webhook URL
    root = tk.Tk()
    root.withdraw()
    webhook = simpledialog.askstring("Webhook", "Enter your Bitrix24 Webhook URL (should end with /rest/):")
    if not webhook: sys.exit("No webhook provided.")
    if not webhook.endswith('/'): webhook += '/'

    # 2. Fetch fields and pipelines
    print("Fetching Bitrix24 fields and pipelines...")
    deal_fields = fetch_fields(webhook, "crm.deal")
    contact_fields = fetch_fields(webhook, "crm.contact")
    pipelines = fetch_pipelines(webhook)
    if not pipelines: sys.exit("No pipelines found.")
    pipeline_names = [f"{p['ID']} - {p['NAME']}" for p in pipelines]
    pipeline_id = simpledialog.askstring("Select Pipeline", "\n".join(pipeline_names) + "\n\nEnter Pipeline ID to use:")
    if not pipeline_id: sys.exit("No pipeline selected.")

    # 3. File selection
    file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
    if not file_path: sys.exit("No file chosen.")
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path, engine="openpyxl")
    columns = list(df.columns)
    print(f"Loaded file with columns: {columns}")

    # 4. Map columns to contact fields
    messagebox.showinfo("Mapping", "Now map Excel columns to Bitrix24 **CONTACT** fields.")
    contact_mapping = mapping_window(columns, contact_fields, "Map Contact Fields")

    # 5. Map columns to deal fields
    messagebox.showinfo("Mapping", "Now map Excel columns to Bitrix24 **DEAL** fields.")
    deal_mapping = mapping_window(columns, deal_fields, "Map Deal Fields")

    # 6. Choose deduplication field
    possible_dedupe_keys = [c for c in contact_mapping.keys()]
    dedupe_field = simpledialog.askstring(
        "Deduplication",
        f"Which Excel column should be used for deduplication when importing contacts?\n\nOptions: {', '.join(possible_dedupe_keys)}\n\n(e.g., Email, Phone, Name)\n\n(If left blank, will use 'Email' if available.)"
    )

    if not dedupe_field or dedupe_field not in contact_mapping:
        # Try default to "Email" (case-insensitive)
        email_candidates = [c for c in contact_mapping.keys() if c.lower() == "email"]
        if email_candidates:
            dedupe_field = email_candidates[0]
            print(f"No valid deduplication field selected. Defaulting to: {dedupe_field}")
        else:
            print("No valid deduplication field selected, and no 'Email' column mapped. Exiting.")
            sys.exit("No valid deduplication field selected.")


    # 7. Choose which DEAL field should be filled with the Bitrix24 Contact ID (usually CONTACT_ID)
    possible_deal_contact_fields = [k for k, v in deal_fields.items() if v.get('type', '').lower() == 'crm_contact' or v.get('title', '').upper().find('CONTACT') >= 0]
    deal_contact_field_choice = simpledialog.askstring(
        "Deal Contact Link",
        f"Which Bitrix24 Deal field should receive the Contact ID?\n\nOptions: {', '.join(possible_deal_contact_fields)}\n\n(Usually CONTACT_ID)"
    )
    if not deal_contact_field_choice or deal_contact_field_choice not in deal_fields:
        deal_contact_field_choice = "CONTACT_ID"  # fallback

    print(f"\nStarting import: deduplication on contact field '{dedupe_field}' and deal field '{deal_contact_field_choice}' for contact link.")

    # 8. Import loop
    for idx, row in df.iterrows():
    # Prepare contact data
        raw_contact_data = {contact_mapping[excel_col]: row[excel_col] for excel_col in contact_mapping if pd.notnull(row[excel_col])}
        contact_data = sanitize_payload(raw_contact_data)
        dedupe_value = row[dedupe_field]
        contact_id, contact_result = None, ""
    try:
        contact_id = find_existing_contact(webhook, contact_mapping[dedupe_field], dedupe_value)
        if not contact_id:
            contact_id = create_contact(webhook, contact_data)
            contact_result = f"Created: {contact_id}" if contact_id else "Create failed"
        else:
            contact_result = f"Found: {contact_id}"
    except Exception as e:
        contact_result = f"Error: {e}"
        contact_id = None

    # Prepare deal data
    raw_deal_data = {deal_mapping[excel_col]: row[excel_col] for excel_col in deal_mapping if pd.notnull(row[excel_col])}
    deal_data = sanitize_payload(raw_deal_data)
    deal_data["CATEGORY_ID"] = pipeline_id
    deal_data[deal_contact_field_choice] = contact_id
    deal_id, deal_result = None, ""
    try:
        if contact_id:
            deal_id = create_deal(webhook, deal_data)
            deal_result = f"Created: {deal_id}" if deal_id else "Create failed"
        else:
            deal_result = "No contact, not created"
    except Exception as e:
        deal_result = f"Error: {e}"
        deal_id = None

    # Write to log
    logwriter.writerow([
        idx+1, dedupe_value,
        repr(contact_data), contact_id, contact_result,
        repr(deal_data), deal_id, deal_result
    ])

    messagebox.showinfo("Done", "Import completed.")

if __name__ == "__main__":
    main()
    logfile.close()
    print(f"\nLog written to: {log_filename}")
