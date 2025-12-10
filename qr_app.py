import os
import qrcode
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, Text
from PIL import Image, ImageTk
from datetime import datetime


# -----------------------------------
# Logging to console box
# -----------------------------------
def log(msg):
    console.insert("end", msg + "\n")
    console.see("end")


# -----------------------------------
# Build text payload
# -----------------------------------
def build_payload(fields):
    return "\n".join(f"{key}: {value}" for key, value in fields.items())


# -----------------------------------
# Open Output Folder
# -----------------------------------
def open_output_folder():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    os.startfile(OUTPUT_FOLDER)
    log(f"Opened Folder: {OUTPUT_FOLDER}")


# -----------------------------------
# Generate QR (Single Entry)
# -----------------------------------
def generate_qr():
    item_code = entry_item_code.get().strip()
    name = entry_name.get().strip()
    department = entry_department.get().strip()
    description = entry_description.get().strip()
    office_address = text_office.get("1.0", "end").strip()
    emergency_name = entry_em_name.get().strip()
    emergency_number = entry_em_number.get().strip()

    if not item_code or not name:
        log("❌ ERROR: Item Code and Name required")
        return

    fields = {
        "Item Code": item_code,
        "Name": name,
        "Department": department,
        "Description": description,
        "Office Address": office_address,
        "Emergency Contact Name": emergency_name,
        "Emergency Contact Number": emergency_number
    }

    payload = build_payload(fields)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c for c in name if c.isalnum() or c in ("_", "-"))
    file_base = f"{item_code}_{safe_name}_{timestamp}"

    png_path = os.path.join(OUTPUT_FOLDER, file_base + ".png")

    qr = qrcode.QRCode(version=None, box_size=10, border=4)
    qr.add_data(payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(png_path)

    # Preview
    preview = Image.open(png_path).resize((240, 240))
    preview_img = ImageTk.PhotoImage(preview)
    qr_preview.config(image=preview_img)
    qr_preview.image = preview_img

    log(f"✔ QR Saved: {png_path}")

    # Show payload in right panel
    payload_box.config(state="normal")
    payload_box.delete("1.0", "end")
    payload_box.insert("1.0", payload)
    payload_box.config(state="disabled")


# -----------------------------------
# Bulk Upload via Excel
# -----------------------------------
def upload_excel():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        log("⚠ No file selected.")
        return

    log(f"Uploading Excel: {file_path}")

    try:
        df = pd.read_excel(file_path)

        required_cols = ["Item Code", "Name"]
        for col in required_cols:
            if col not in df.columns:
                log(f"❌ Missing column: {col}")
                return

        count = 0
        for _, row in df.iterrows():
            item_code = str(row.get("Item Code", "")).strip()
            name = str(row.get("Name", "")).strip()

            # category handle
            category = row.get("Category", "")
            if pd.notna(category) and category != "":
                item_code = f"{item_code}_{category}"

            fields = {
                "Item Code": item_code,
                "Name": name,
                "Department": str(row.get("Department", "")),
                "Description": str(row.get("Description", "")),
                "Office Address": str(row.get("Office Address", "")),
                "Emergency Contact Name": str(row.get("Emergency Contact Name", "")),
                "Emergency Contact Number": str(row.get("Emergency Contact Number", ""))
            }

            payload = build_payload(fields)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = "".join(c for c in name if c.isalnum() or c in ("_", "-"))
            file_base = f"{item_code}_{safe_name}_{timestamp}"

            png_path = os.path.join(OUTPUT_FOLDER, file_base + ".png")

            qr = qrcode.QRCode(version=None, box_size=10, border=4)
            qr.add_data(payload)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            img.save(png_path)

            count += 1

        log(f"✔ Bulk QR Complete: {count} items")

    except Exception as e:
        log("❌ Error: " + str(e))


# -----------------------------------
# UI with ttkbootstrap
# -----------------------------------
# app = ttk.Window(title="Bridgelab QR Generator", themename="minty", size=(1000, 650))
app = ttk.Window(title="Bridgelab QR Generator", themename="minty")
app.state("zoomed")     # Windows Full Screen

OUTPUT_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "Bridgelab_QR")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# Left Panel
left = ttk.Frame(app, padding=15)
left.pack(side="left", fill="y")

ttk.Label(left, text="Item Code *").pack(anchor="w")
entry_item_code = ttk.Entry(left)
entry_item_code.pack(fill="x")

ttk.Label(left, text="Name *").pack(anchor="w")
entry_name = ttk.Entry(left)
entry_name.pack(fill="x")

ttk.Label(left, text="Department").pack(anchor="w")
entry_department = ttk.Entry(left)
entry_department.pack(fill="x")

ttk.Label(left, text="Description").pack(anchor="w")
entry_description = ttk.Entry(left)
entry_description.pack(fill="x")

ttk.Label(left, text="Office Address").pack(anchor="w")
text_office = Text(left, height=4)
text_office.pack(fill="x")

ttk.Label(left, text="Emergency Contact Name").pack(anchor="w")
entry_em_name = ttk.Entry(left)
entry_em_name.pack(fill="x")

ttk.Label(left, text="Emergency Contact Number").pack(anchor="w")
entry_em_number = ttk.Entry(left)
entry_em_number.pack(fill="x")

ttk.Button(left, text="Generate QR", bootstyle=PRIMARY, command=generate_qr).pack(fill="x", pady=5)
ttk.Button(left, text="Upload Excel (Bulk)", bootstyle=SUCCESS, command=upload_excel).pack(fill="x", pady=5)
ttk.Button(left, text="Open Output Folder", bootstyle=INFO, command=open_output_folder).pack(fill="x", pady=5)

# Right Section
right = ttk.Frame(app, padding=15)
right.pack(side="right", fill="both", expand=True)

ttk.Label(right, text="QR Preview", font=("Arial", 14, "bold")).pack()
qr_preview = ttk.Label(right)
qr_preview.pack(pady=10)

ttk.Label(right, text="QR Payload", font=("Arial", 12, "bold")).pack()
payload_box = Text(right, height=10)
payload_box.pack(fill="x")
payload_box.config(state="disabled")

ttk.Label(right, text="Console Log", font=("Arial", 12, "bold")).pack()
console = Text(right, height=10)
console.pack(fill="both", expand=True)

log("Ready...")

app.mainloop()

# import os
# import qrcode
# import pandas as pd
# from PIL import Image, ImageTk
# import tkinter as tk
# from tkinter import messagebox, filedialog, scrolledtext
# from datetime import datetime

# # ----------------------------
# # Helper: build QR payload (no timestamp inside)
# # ----------------------------
# def build_payload(fields):
#     lines = []
#     for label, value in fields.items():
#         lines.append(f"{label}: {value}")
#     return "\n".join(lines)

# # ----------------------------
# # Open the output folder
# # ----------------------------
# def open_output_folder():
#     desktop = os.path.join(os.path.expanduser("~"), "Desktop")
#     output_folder = os.path.join(desktop, "Bridgelab_QR")
#     os.makedirs(output_folder, exist_ok=True)
#     os.startfile(output_folder)

# # ----------------------------
# # Main QR generator (single entry)
# # ----------------------------
# def generate_qr(save_txt=True):
#     item_code = entry_item_code.get().strip()
#     name = entry_name.get().strip()
#     department = entry_department.get().strip()
#     description = entry_description.get().strip()
#     office_address = text_office_address.get("1.0", "end").strip()
#     emergency_name = entry_emergency_name.get().strip()
#     emergency_number = entry_emergency_number.get().strip()

#     if not item_code or not name:
#         messagebox.showerror("Missing Fields", "Item Code and Name are required.")
#         return

#     fields = {
#         "Item Code": item_code,
#         "Name": name,
#         "Department": department,
#         "Description": description,
#         "Office Address": office_address,
#         "Emergency Contact Name": emergency_name,
#         "Emergency Contact Number": emergency_number
#     }

#     payload = build_payload(fields)

#     desktop = os.path.join(os.path.expanduser("~"), "Desktop")
#     output_folder = os.path.join(desktop, "Bridgelab_QR")
#     os.makedirs(output_folder, exist_ok=True)

#     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#     safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).rstrip()
#     filename_base = f"{item_code}_{safe_name}_{timestamp}"

#     png_path = os.path.join(output_folder, filename_base + ".png")
#     txt_path = os.path.join(output_folder, filename_base + ".txt")

#     qr = qrcode.QRCode(version=None, box_size=10, border=4)
#     qr.add_data(payload)
#     qr.make(fit=True)
#     img = qr.make_image(fill_color="black", back_color="white")
#     img.save(png_path)

#     if save_txt:
#         with open(txt_path, "w", encoding="utf-8") as f:
#             f.write(payload)

#     preview_img = Image.open(png_path).resize((240, 240))
#     tk_img = ImageTk.PhotoImage(preview_img)
#     qr_preview_label.config(image=tk_img)
#     qr_preview_label.image = tk_img

#     payload_box.config(state="normal")
#     payload_box.delete("1.0", "end")
#     payload_box.insert("1.0", payload)
#     payload_box.config(state="disabled")

#     messagebox.showinfo("Saved", f"QR Saved:\n{png_path}")

# # ----------------------------
# # Excel Upload & Bulk QR Creator
# # ----------------------------
# def upload_excel():
#     file_path = filedialog.askopenfilename(
#         title="Select Excel File",
#         filetypes=[("Excel Files", "*.xlsx")]
#     )

#     if not file_path:
#         return

#     try:
#         df = pd.read_excel(file_path)

#         # Required minimum columns
#         required = ["Item Code", "Name"]
#         for col in required:
#             if col not in df.columns:
#                 messagebox.showerror("Error", f"Missing required column: {col}")
#                 return

#         desktop = os.path.join(os.path.expanduser("~"), "Desktop")
#         output_folder = os.path.join(desktop, "Bridgelab_QR")
#         os.makedirs(output_folder, exist_ok=True)

#         count = 0
#         for index, row in df.iterrows():
#             item_code = str(row.get("Item Code", "")).strip()
#             name = str(row.get("Name", "")).strip()

#             category = row.get("Category", "")
#             if pd.notna(category):
#                 item_code = f"{item_code}_{category}"

#             fields = {
#                 "Item Code": item_code,
#                 "Name": name,
#                 "Department": str(row.get("Department", "")),
#                 "Description": str(row.get("Description", "")),
#                 "Office Address": str(row.get("Office Address", "")),
#                 "Emergency Contact Name": str(row.get("Emergency Contact Name", "")),
#                 "Emergency Contact Number": str(row.get("Emergency Contact Number", ""))
#             }

#             payload = build_payload(fields)

#             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#             safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).rstrip()
#             filename_base = f"{item_code}_{safe_name}_{timestamp}"

#             png_path = os.path.join(output_folder, filename_base + ".png")
#             txt_path = os.path.join(output_folder, filename_base + ".txt")

#             qr = qrcode.QRCode(version=None, box_size=10, border=4)
#             qr.add_data(payload)
#             qr.make(fit=True)
#             img = qr.make_image(fill_color="black", back_color="white")
#             img.save(png_path)

#             with open(txt_path, "w", encoding="utf-8") as f:
#                 f.write(payload)

#             count += 1

#         messagebox.showinfo("Success", f"{count} QR Codes generated from Excel.")

#     except Exception as e:
#         messagebox.showerror("Error", str(e))

# # ----------------------------
# # GUI Layout
# # ----------------------------
# root = tk.Tk()
# root.title("Bridgelab - QR Generator (Plain Text)")
# root.geometry("900x650")
# root.resizable(False, False)

# left = tk.Frame(root, padx=12, pady=12)
# left.pack(side="left", fill="y")

# right = tk.Frame(root, padx=12, pady=12)
# right.pack(side="right", fill="both", expand=True)

# # Inputs
# tk.Label(left, text="Item Code *").pack(anchor="w")
# entry_item_code = tk.Entry(left, width=32)
# entry_item_code.pack()

# tk.Label(left, text="Name *").pack(anchor="w")
# entry_name = tk.Entry(left, width=32)
# entry_name.pack()

# tk.Label(left, text="Department").pack(anchor="w")
# entry_department = tk.Entry(left, width=32)
# entry_department.pack()

# tk.Label(left, text="Description").pack(anchor="w")
# entry_description = tk.Entry(left, width=32)
# entry_description.pack()

# tk.Label(left, text="Office Address").pack(anchor="w")
# text_office_address = tk.Text(left, width=32, height=4)
# text_office_address.pack()

# tk.Label(left, text="Emergency Contact Name").pack(anchor="w")
# entry_emergency_name = tk.Entry(left, width=32)
# entry_emergency_name.pack()

# tk.Label(left, text="Emergency Contact Number").pack(anchor="w")
# entry_emergency_number = tk.Entry(left, width=32)
# entry_emergency_number.pack()

# tk.Button(left, text="Generate QR", width=28, bg="#1976D2", fg="white",
#           command=generate_qr).pack(pady=8)

# tk.Button(left, text="Upload Excel (Bulk QR)", width=28, bg="#43A047", fg="white",
#           command=upload_excel).pack(pady=4)

# tk.Button(left, text="Open Output Folder", width=28, bg="#6A1B9A", fg="white",
#           command=open_output_folder).pack(pady=4)

# tk.Label(left, text="* Required Fields", fg="gray").pack(anchor="w", pady=6)

# # Right section
# tk.Label(right, text="QR Preview", font=("Arial", 12, "bold")).pack()
# qr_preview_label = tk.Label(right, width=240, height=240, relief="groove")
# qr_preview_label.pack(pady=8)

# tk.Label(right, text="QR Payload", font=("Arial", 12, "bold")).pack()
# payload_box = scrolledtext.ScrolledText(right, width=50, height=12)
# payload_box.pack()
# payload_box.config(state="disabled")

# root.mainloop()

