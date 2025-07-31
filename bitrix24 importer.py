import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import requests
import pandas as pd
import sys

def get_webhook():
    return simpledialog.askstring("Webhook", "Enter your Bitrix24 Webhook URL (should end with /rest/):")

def fetch_fields(webhook, entity):
    resp = requests.get(f"{webhook}{entity}.fields.json")
    return resp.json()["result"]

def fetch_pipelines(webhook):
    resp = requests.get(f"{webhook}crm.dealcategory.list.json")
    return resp.json()["result"]

def pick_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    return file_path

def select_pipeline(pipelines):
    pipeline_names = [f"{p['ID']} - {p['NAME']}" for p in pipelines]
    selection = simpledialog.askstring("Select Pipeline", "\n".join(pipeline_names) + "\n\nEnter Pipeline ID to use:")
    return selection

def column_mapping(columns, b24_fields, title):
    import tkinter as tk
    from tkinter import ttk

    mapping = {}

    root = tk.Tk()
    root.title(title)

    # Set window size and make it resizable
    root.geometry("600x500")
    root.resizable(True, True)

    # Add a canvas for scrolling
    canvas = tk.Canvas(root)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Add a vertical scrollbar to the canvas
    vscrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=canvas.yview)
    vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Configure canvas
    canvas.configure(yscrollcommand=vscrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Create a frame inside the canvas
    frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    dropdown_vars = []
    dropdowns = []

    # Sort Bitrix fields alphabetically by title
    b24_choices = [""] + [f"{fid}: {fdata['title']}" for fid, fdata in sorted(b24_fields.items(), key=lambda x: x[1]['title'].lower())]
    b24_ids = [""] + list(sorted(b24_fields.keys(), key=lambda k: b24_fields[k]['title'].lower()))

    # Header
    tk.Label(frame, text="Excel Column", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tk.Label(frame, text="Bitrix24 Field", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=5)

    # Add dropdowns for each column
    for i, col in enumerate(columns):
        tk.Label(frame, text=col, anchor="w").grid(row=i+1, column=0, sticky="w", padx=5, pady=3)
        var = tk.StringVar()
        dropdown = ttk.Combobox(frame, textvariable=var, values=b24_choices, state="readonly", width=45)
        dropdown.grid(row=i+1, column=1, sticky="w", padx=5, pady=3)
        dropdown_vars.append(var)
        dropdowns.append(dropdown)

    def submit():
        for i, col in enumerate(columns):
            selected = dropdown_vars[i].get()
            if selected:
                try:
                    # Extract field ID from "id: title"
                    fid = selected.split(":")[0]
                    if fid in b24_fields:
                        mapping[col] = fid
                except Exception:
                    continue
        root.destroy()

    submit_btn = tk.Button(frame, text="Submit", command=submit)
    submit_btn.grid(row=len(columns) + 2, column=0, columnspan=2, pady=20)

    # Enable mousewheel scrolling (Windows & Mac)
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    # For Linux (needs '<Button-4>' and '<Button-5>' events)
    canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
    canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))

    root.mainloop()
    return mapping


def create_contact(webhook, contact_payload):
    r = requests.post(f"{webhook}crm.contact.add.json", json={"fields": contact_payload, "params": {"REGISTER_SONET_EVENT": "N"}})
    data = r.json()
    if 'result' in data:
        return data['result']
    else:
        print("Contact creation failed:", data)
        return None

def create_deal(webhook, deal_payload):
    r = requests.post(f"{webhook}crm.deal.add.json", json={"fields": deal_payload, "params": {"REGISTER_SONET_EVENT": "N"}})
    data = r.json()
    if 'result' in data:
        return data['result']
    else:
        print("Deal creation failed:", data)
        return None

def main():
    # 1. Ask for webhook
    root = tk.Tk()
    root.withdraw()
    webhook = get_webhook()
    if not webhook:
        sys.exit("No webhook provided.")
    if not webhook.endswith('/'):
        webhook += '/'

    # 2. Fetch fields and pipelines
    deal_fields = fetch_fields(webhook, "crm.deal")
    contact_fields = fetch_fields(webhook, "crm.contact")
    pipelines = fetch_pipelines(webhook)
    if not pipelines:
        sys.exit("No pipelines found.")
    pipeline_id = select_pipeline(pipelines)
    if not pipeline_id:
        sys.exit("No pipeline selected.")

    # 3. Pick Excel file
    file_path = pick_file()
    if not file_path:
        sys.exit("No file chosen.")
    df = pd.read_excel(file_path)
    columns = list(df.columns)

    # 4. Pair columns to Bitrix24 fields
    deal_mapping = column_mapping(columns, deal_fields, "Map Deal Fields")
    contact_mapping = column_mapping(columns, contact_fields, "Map Contact Fields")

    # 5. Import rows
    for idx, row in df.iterrows():
        # 5.1 Prepare contact fields
        contact_data = {b24_field: row[excel_col] for excel_col, b24_field in contact_mapping.items() if pd.notnull(row[excel_col])}
        contact_id = None
        if contact_data:
            contact_id = create_contact(webhook, contact_data)
        # 5.2 Prepare deal fields
        deal_data = {b24_field: row[excel_col] for excel_col, b24_field in deal_mapping.items() if pd.notnull(row[excel_col])}
        deal_data["CATEGORY_ID"] = pipeline_id
        if contact_id:
            deal_data["CONTACT_ID"] = contact_id
        deal_id = create_deal(webhook, deal_data)
        print(f"Imported row {idx+1}: Contact ID {contact_id} - Deal ID {deal_id}")

    messagebox.showinfo("Done", "Import completed.")

if __name__ == "__main__":
    main()
