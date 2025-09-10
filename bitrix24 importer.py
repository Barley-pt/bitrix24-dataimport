import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from collections import defaultdict
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

# --- GUI Mapping Window with Multi-field Type Selection ---
def mapping_window(columns, b24_fields, title):
    import re
    mapping = {}
    window = tk.Toplevel()
    window.title(title)
    window.geometry("900x500")
    window.minsize(600, 300)
    window.resizable(True, True)
    canvas = tk.Canvas(window)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    vscrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=canvas.yview)
    vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.configure(yscrollcommand=vscrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    # Prepare Bitrix24 field choices and map for value types
    b24_choices = [""] + [field_label(fid, fdata) for fid, fdata in sorted(b24_fields.items(), key=lambda x: field_label(x[0], x[1]).lower())]
    b24_id_map = {field_label(fid, fdata): fid for fid, fdata in b24_fields.items()}
    multifields = {"PHONE", "EMAIL", "IM"}  # can add more if needed

    # Common Bitrix24 types (add more if needed)
    value_types = {
        "PHONE": ["WORK", "MOBILE", "HOME", "FAX", "OTHER"],
        "EMAIL": ["WORK", "HOME", "OTHER"],
        "IM": ["SKYPE", "ICQ", "MSN", "JABBER", "OTHER"],
    }

    tk.Label(frame, text="Excel Column", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tk.Label(frame, text="Bitrix24 Field", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=5)
    tk.Label(frame, text="Value Type", font=("Arial", 12, "bold")).grid(row=0, column=2, sticky="w", padx=5, pady=5)

    dropdown_vars = []
    type_vars = []
    type_boxes = []

    for i, col in enumerate(columns):
        tk.Label(frame, text=col, anchor="w").grid(row=i+1, column=0, sticky="w", padx=5, pady=3)
        field_var = tk.StringVar()
        type_var = tk.StringVar()
        dropdown = ttk.Combobox(frame, textvariable=field_var, values=b24_choices, state="readonly", width=55)
        dropdown.grid(row=i+1, column=1, sticky="w", padx=5, pady=3)
        type_box = ttk.Combobox(frame, textvariable=type_var, values=[], state="readonly", width=14)
        type_box.grid(row=i+1, column=2, sticky="w", padx=5, pady=3)
        dropdown_vars.append(field_var)
        type_vars.append(type_var)
        type_boxes.append(type_box)

        # When the Bitrix24 field is changed, show types if needed
        def on_field_select(event, idx=i):
            field = dropdown_vars[idx].get()
            fid = b24_id_map.get(field)
            type_box = type_boxes[idx]
            show_types = False
            if fid and fid.upper() in multifields:
                show_types = True
            if show_types:
                type_box["values"] = value_types.get(fid.upper(), ["WORK", "HOME", "OTHER"])
                # Try to auto-select based on Excel col name (e.g., if col contains "mobile")
                col_lower = columns[idx].lower()
                auto_type = ""
                if "mobile" in col_lower:
                    auto_type = "MOBILE"
                elif "work" in col_lower:
                    auto_type = "WORK"
                elif "home" in col_lower:
                    auto_type = "HOME"
                elif "fax" in col_lower:
                    auto_type = "FAX"
                elif "skype" in col_lower:
                    auto_type = "SKYPE"
                if auto_type and auto_type in type_box["values"]:
                    type_vars[idx].set(auto_type)
                else:
                    type_vars[idx].set(type_box["values"][0])
                type_box["state"] = "readonly"
            else:
                type_box.set("")
                type_box["values"] = []
                type_box["state"] = "disabled"
        dropdown.bind("<<ComboboxSelected>>", on_field_select)
        type_box["state"] = "disabled"

    def submit():
        for i, col in enumerate(columns):
            selected_field = dropdown_vars[i].get()
            selected_type = type_vars[i].get()
            fid = b24_id_map.get(selected_field)
            # If type is set, return as tuple; else None for value type
            if fid:
                mapping[col] = (fid, selected_type if selected_type else None)
        window.destroy()
    submit_btn = tk.Button(frame, text="Submit", command=submit)
    submit_btn.grid(row=len(columns) + 2, column=0, columnspan=3, pady=20)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
    canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))
    window.grab_set()
    window.wait_window()
    return mapping

# --- Build payloads supporting multi-fields for Bitrix24 ---
def build_multifield_payload(row, mapping):
    """
    Build Bitrix24 contact/deal payload from mapping:
    mapping: {Excel column: (Bitrix24 field, value_type or None)}
    Returns: dict ready to send to Bitrix24 API.
    """
    from collections import defaultdict
    import pandas as pd
    import datetime

    multifields = {"PHONE", "EMAIL", "IM"}
    multifield_payloads = defaultdict(list)
    simple_payload = {}
    for excel_col, (fid, vtype) in mapping.items():
        value = row[excel_col]
        if pd.isnull(value) or value == '':
            continue
        # Dates/timestamps to string
        if isinstance(value, pd.Timestamp):
            value = value.strftime("%Y-%m-%d")
        elif isinstance(value, (datetime.date, datetime.datetime)):
            value = value.isoformat()

        # Multi-field logic
        if fid in multifields and vtype:
            # Accept split values by , ; or |
            vals = [v.strip() for v in str(value).replace(";",",").replace("|",",").split(",") if v.strip()]
            for val in vals:
                multifield_payloads[fid].append({"VALUE": val, "VALUE_TYPE": vtype})
        else:
            simple_payload[fid] = value
    # Merge
    payload = simple_payload.copy()
    payload.update(multifield_payloads)
    return payload

# --- Fetch fields/pipelines from Bitrix24 ---
def fetch_fields(webhook, entity):
    resp = requests.get(f"{webhook}{entity}.fields.json")
    resp.raise_for_status()
    return resp.json()["result"]

def fetch_pipelines(webhook):
    resp = requests.get(f"{webhook}crm.dealcategory.list.json")
    resp.raise_for_status()
    return resp.json()["result"]

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
    webhook = simpledialog.askstring("Webhook", "Enter your Bitrix24 Webhook URL:")
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
    import time

    for idx, row in df.iterrows():
        print(f"\n[{idx+1}/{len(df)}] Importing row...")

        # Prepare contact data
        contact_data = build_multifield_payload(row, contact_mapping)
        print("  Contact payload:")
        for k, v in contact_data.items():
            print(f"    {k}: {v}")

        dedupe_value = row[dedupe_field]
        contact_id, contact_result = None, ""
        try:
            contact_id = find_existing_contact(webhook, contact_mapping[dedupe_field][0], dedupe_value)
            time.sleep(0.5)  # Bitrix24 rate limit
            if not contact_id:
                contact_id = create_contact(webhook, contact_data)
                contact_result = f"Created: {contact_id}" if contact_id else "Create failed"
                print(f"  ➔ Contact {contact_result}")
            else:
                contact_result = f"Found: {contact_id}"
                print(f"  ➔ Existing contact found: {contact_id}")
        except Exception as e:
            contact_result = f"Error: {e}"
            contact_id = None
            print(f"  ➔ Error searching/creating contact: {e}")

        # Prepare deal data
        deal_data = build_multifield_payload(row, deal_mapping)
        deal_data["CATEGORY_ID"] = pipeline_id
        deal_data[deal_contact_field_choice] = contact_id
        print("  Deal payload:")
        for k, v in deal_data.items():
            print(f"    {k}: {v}")
        deal_id, deal_result = None, ""
        try:
            if contact_id:
                deal_id = create_deal(webhook, deal_data)
                time.sleep(0.5)  # Bitrix24 rate limit
                deal_result = f"Created: {deal_id}" if deal_id else "Create failed"
                print(f"  ➔ Deal {deal_result}")
            else:
                deal_result = "No contact, not created"
                print(f"  ➔ Deal not created (no contact)")
        except Exception as e:
            deal_result = f"Error: {e}"
            deal_id = None
            print(f"  ➔ Error creating deal: {e}")

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
