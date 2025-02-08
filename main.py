import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from PIL import Image, ImageTk
import openpyxl
import ttkbootstrap as tb


# Excel file path
EXCEL_FILE = "record.xlsx"


# Ensure Excel file exists with headers
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Name", "Phone", "Photo", "Signature", "Unique ID", "Purpose", "Service", "Vehicle Number", "Email"])
    workbook.save(EXCEL_FILE)


# Function to append data to Excel
def submit_form():
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active

    data = [
        entry_name.get(),
        entry_phone.get(),
        photo_path.get(),
        signature_path.get(),
        entry_unique_id.get(),
        purpose_var.get(),
        service_var.get(),
        entry_vehicle_number.get(),
        entry_email.get(),
    ]

    sheet.append(data)
    workbook.save(EXCEL_FILE)
    messagebox.showinfo("Success", "Data appended to Excel successfully!")
    clear_form()

# Function to browse files
def select_file(entry_var):
    file_path = filedialog.askopenfilename()
    entry_var.set(file_path)

# Function to clear form fields
def clear_form():
    entry_name.delete(0, tk.END)
    entry_phone.delete(0, tk.END)
    photo_path.set("")
    signature_path.set("")
    entry_unique_id.delete(0, tk.END)
    purpose_var.set("")
    service_var.set("")
    entry_vehicle_number.delete(0, tk.END)
    entry_email.delete(0, tk.END)

# Create the main window with ttkbootstrap theme
root = tb.Window(themename="flatly")
root.title("उप प्रादेशिक परिवहन कार्यालय, जालना")
root.geometry("900x900")

# Load and set the background image
bg_image = Image.open("RTO.png")  # Replace with your image
bg_image = bg_image.resize((900, 900))
bg_photo = ImageTk.PhotoImage(bg_image)

bg_label = tk.Label(root, image=bg_photo)
bg_label.place(relwidth=1, relheight=1)

# Frame to hold the form with semi-transparent background
frame = tk.Frame(root, bg="#ffffff", bd=5)
frame.place(relx=0.5, rely=0.55, anchor="center", width=600, height=700)

# Load and place the logos
logo1 = Image.open("logo1.png")  # Replace with actual logo file
logo1 = logo1.resize((80, 80))
logo1 = ImageTk.PhotoImage(logo1)

logo2 = Image.open("logo2.png")  # Replace with actual logo file
logo2 = logo2.resize((80, 80))
logo2 = ImageTk.PhotoImage(logo2)

logo3 = Image.open("logo3.png")  # Replace with actual logo file
logo3 = logo3.resize((80, 80))
logo3 = ImageTk.PhotoImage(logo3)

tk.Label(root, image=logo1, bg="white").place(x=50, y=20)  # Left
tk.Label(root, image=logo2, bg="white").place(x=400, y=10)  # Center
tk.Label(root, image=logo3, bg="white").place(x=750, y=20)  # Right

# Title
ttk.Label(frame, text="Customer Form", font=("Arial", 20, "bold"), background="white").pack(pady=10)

# Form Fields
fields = [
    ("Name", "entry_name"),
    ("Phone", "entry_phone"),
    ("Unique ID (Optional)", "entry_unique_id"),
    ("Vehicle Number", "entry_vehicle_number"),
    ("Email", "entry_email"),
]

entries = {}
for label_text, var_name in fields:
    frame_sub = ttk.Frame(frame, style="TFrame")
    frame_sub.pack(pady=10, fill="x", padx=40)
    
    ttk.Label(frame_sub, text=label_text, font=("Arial", 14), background="white").pack(side="left")
    entry = ttk.Entry(frame_sub, width=40, font=("Arial", 12))
    entry.pack(side="right")
    entries[var_name] = entry

entry_name, entry_phone, entry_unique_id, entry_vehicle_number, entry_email = (
    entries["entry_name"],
    entries["entry_phone"],
    entries["entry_unique_id"],
    entries["entry_vehicle_number"],
    entries["entry_email"],
)

# Photo Selection
photo_frame = ttk.Frame(frame)
photo_frame.pack(pady=10, fill="x", padx=40)

photo_path = tk.StringVar()
ttk.Label(photo_frame, text="Photo:", font=("Arial", 14), background="white").pack(side="left")
ttk.Entry(photo_frame, textvariable=photo_path, width=30, font=("Arial", 12)).pack(side="left", padx=5)
ttk.Button(photo_frame, text="Browse", command=lambda: select_file(photo_path), bootstyle="primary").pack(side="right")

# Signature Selection
signature_frame = ttk.Frame(frame)
signature_frame.pack(pady=10, fill="x", padx=40)

signature_path = tk.StringVar()
ttk.Label(signature_frame, text="Signature:", font=("Arial", 14), background="white").pack(side="left")
ttk.Entry(signature_frame, textvariable=signature_path, width=30, font=("Arial", 12)).pack(side="left", padx=5)
ttk.Button(signature_frame, text="Browse", command=lambda: select_file(signature_path), bootstyle="primary").pack(side="right")

# Purpose of Visit
purpose_frame = ttk.Frame(frame)
purpose_frame.pack(pady=10, fill="x", padx=40)

ttk.Label(purpose_frame, text="Purpose:", font=("Arial", 14), background="white").pack(side="left")
purpose_var = tk.StringVar()
purpose_dropdown = ttk.Combobox(purpose_frame, textvariable=purpose_var, width=35, state="readonly", font=("Arial", 12))
purpose_dropdown["values"] = ("Vehicle Related", "License Related")
purpose_dropdown.pack(side="right")

# Service Selection
service_frame = ttk.Frame(frame)
service_frame.pack(pady=10, fill="x", padx=40)

ttk.Label(service_frame, text="Service:", font=("Arial", 14), background="white").pack(side="left")
service_var = tk.StringVar()
service_dropdown = ttk.Combobox(service_frame, textvariable=service_var, width=35, state="readonly", font=("Arial", 12))
service_dropdown["values"] = (
    "Transfer of ownership - seller",
    "Transfer of ownership - purchaser",
    "Issue duplicate RC",
    "Hypothecation termination",
    "Hypothecation addition",
    "MV permit related",
    "MDL renewal",
    "PSV driver Badge renewal",
    "International driving permit",
    "Application for New MDL",
    "Learning license related",
    "Other (Specify)",
)
service_dropdown.pack(side="right")

# Buttons
btn_frame = ttk.Frame(frame, style="TFrame")
btn_frame.pack(pady=20)

ttk.Button(btn_frame, text="Submit", command=submit_form, bootstyle="success", style="TButton").pack(side="left", padx=20)
ttk.Button(btn_frame, text="Clear", command=clear_form, bootstyle="danger", style="TButton").pack(side="right", padx=20)

# Run the application
root.mainloop()
















