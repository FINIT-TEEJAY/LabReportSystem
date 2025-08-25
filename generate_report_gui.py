# ========================== Imports ==========================
import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from PIL import Image, ImageTk
import re
from datetime import datetime

# ========================== Utility Function ==========================
# Function to alter CRP and RF values for specific report formatting
def get_altered_result(test_name: str, result):
    if test_name and (result or result == 0):
        try:
            result_float = float(result)
        except ValueError:
            return result  # Return original if not a number

        if test_name == "CRP" and result_float < 6:
            return "&lt;6"
        elif test_name == "RF" and (result_float < 10 or result_float == 0):
            return "&lt;10"
    return result

# ========================== Data Loading ==========================
# Load test packages
packages_df = pd.read_excel("test_packages.xlsx")
package_names = sorted(packages_df['package'].unique())

# Load lab results
lab_df = pd.read_excel("lab_data.xlsx")
barcodes_list = sorted(lab_df['barcode'].astype(str).unique())

# Load patient data
patient_df = pd.read_excel("patient_data.xlsx")

# ========================== Event Handlers ==========================
# Function to auto-fill patient details from barcode
def autofill_patient_info(event=None):
    barcode = combo_barcode.get().strip()
    if not barcode:
        return

    lab_match = lab_df[lab_df['barcode'].astype(str).str.strip() == barcode]
    if lab_match.empty:
        messagebox.showerror("Not Found", f"No lab data found for barcode: {barcode}")
        return

    patient_name = lab_match.iloc[0]['patient_name'].strip()
    patient_match = patient_df[patient_df['patient_name'].str.strip() == patient_name]
    if patient_match.empty:
        messagebox.showerror("Not Found", f"No patient details found for name: {patient_name}")
        return

    row = patient_match.iloc[0]

    # Fill form fields
    entry_name.delete(0, tk.END)
    entry_name.insert(0, patient_name)

    entry_gender.delete(0, tk.END)
    entry_gender.insert(0, row.get('gender', ''))

    dob_raw = row.get('dob', '')

    # Try multiple formats for DOB
    dob_formatted = str(dob_raw)  # fallback
    if dob_raw:
        if isinstance(dob_raw, datetime):
            dob_formatted = dob_raw.strftime("%d-%m-%Y")
        else:
        # Try multiple formats
            parsed = False
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
                try:
                    dob_formatted = datetime.strptime(str(dob_raw), fmt).strftime("%d-%m-%Y")
                    parsed = True
                    break
                except ValueError:
                    continue
        if not parsed:
            print(f"Could not parse DOB: {dob_raw}")

# Insert into DOB entry
    entry_dob.delete(0, tk.END)
    entry_dob.insert(0, dob_formatted)

    current_date = datetime.now().strftime("%d-%m-%Y")
    
    entry_collected_on.delete(0, tk.END)
    entry_collected_on.insert(0, current_date)
    
    entry_recieved_on.delete(0, tk.END)
    entry_recieved_on.insert(0, current_date)
    
    entry_report_on.delete(0, tk.END)
    entry_report_on.insert(0, current_date)

    entry_hospital_number.delete(0, tk.END)
    entry_hospital_number.insert(0,  "-")

    entry_address.delete(0, tk.END)
    entry_address.insert(0, row.get('address', ''))

    combo_package.set(row.get('package', ''))

# Function to clear all fields
def refresh_form():
    combo_barcode.set("")
    entry_name.delete(0, tk.END)
    entry_gender.delete(0, tk.END)
    entry_dob.delete(0, tk.END)
    entry_collected_on.delete(0, tk.END)
    entry_recieved_on.delete(0, tk.END)
    entry_report_on.delete(0, tk.END)
    entry_hospital_number.delete(0, tk.END)
    entry_address.delete(0, tk.END)
    combo_package.set("")

# ========================== Report Generation ==========================
# Function to generate the report using docx template
def generate_report():
    barcode = combo_barcode.get().strip()
    name = entry_name.get().strip()
    gender = entry_gender.get().strip()
    dob = entry_dob.get().strip()
    package = combo_package.get().strip()
    package_cleaned = package.lower().strip()
    collected_on = entry_collected_on.get().strip()
    report_on = entry_report_on.get().strip()
    recieved_on = entry_recieved_on.get().strip()
    hospital_number = entry_hospital_number.get().strip()
    address = entry_address.get().strip()

    # Validate fields
    if not all([barcode, dob, package, collected_on, report_on, recieved_on, hospital_number, address]):
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return

    # Normalize test names
    lab_df['test_name_normalized'] = lab_df['test_name'].str.strip().str.lower()
    

    # Get required tests from the package
    required_tests = packages_df[
        packages_df['package'].str.strip().str.lower() == package_cleaned
    ]['test_name'].dropna().str.strip().str.lower().tolist()

    if not required_tests:
        messagebox.showerror("Package Error", f"No tests found for package: {package}")
        return

    # Filter test results for this barcode and required tests
    df = lab_df[
        (lab_df['barcode'].astype(str).str.strip() == barcode) &
        (lab_df['test_name_normalized'].isin(required_tests))
    ]
        # Filter test results for this barcode and required tests
    df = lab_df[
        (lab_df['barcode'].astype(str).str.strip() == barcode) &
        (lab_df['test_name_normalized'].isin(required_tests))
    ]

    # Check and show info about any missing tests
    found_tests = df['test_name_normalized'].unique().tolist()
    missing_tests = [test for test in required_tests if test not in found_tests]

    if missing_tests:
        formatted_missing = ", ".join(test.title() for test in missing_tests)
        messagebox.showinfo(
            "Missing Tests",
            f"{len(missing_tests)} test(s) not found for selected package: {package}\n\nMissing tests:\n{formatted_missing}"
        )

    if df.empty:
        messagebox.showerror("No Data", "No matching test results found.")
        return

    # Prepare results for the report
    results = []
    named_results = {}

    for _, row in df.iterrows():
        test_name = row['test_name'].strip()
        value = row['value']
        unit = row['unit']
        altered_value = get_altered_result(test_name, value)
        test_key = re.sub(r'[^a-zA-Z0-9_]', '_', test_name.lower())
        named_results[f"{test_key}_value"] = altered_value
        named_results[f"{test_key}_unit"] = unit
        results.append({"name": test_name, "value": altered_value, "unit": unit})

    # Context for docx template
    context = {
        "barcode": barcode,
        "patient_name": name,
        "gender": gender,
        "dob": dob,
        "test_package": package,
        "results": results,
        "collected_on": collected_on,
        "report_on": report_on,
        "recieved_on": recieved_on,
        "hospital_number": hospital_number,
        "address": address,
        **named_results
    }

    # Template and output paths
    template_file = f"templates/template_{package}.docx"
    if not os.path.exists(template_file):
        messagebox.showerror("Template Error", f"Template not found: {template_file}")
        return

    doc = DocxTemplate(template_file)
    doc.render(context)

    os.makedirs("reports", exist_ok=True)
    output_docx = f"reports/report_{barcode}.docx"
    output_pdf = output_docx.replace(".docx", ".pdf")
    doc.save(output_docx)

    try:
        convert(output_docx, output_pdf)
    except:
        messagebox.showwarning("PDF Warning", "PDF conversion failed. Word report saved.")

    msg = f"✅ Word report created: {output_docx}"
    if os.path.exists(output_pdf):
        msg += f"\n📄 PDF created: {output_pdf}"
    messagebox.showinfo("Success", msg)

# ========================== GUI Setup ==========================
# Main application window
root = tk.Tk()
root.title("MediLab Report Generator")

# Logo display
try:
    pil_image = Image.open("logo.png")
    resized_image = pil_image.resize((300, 120))
    logo = ImageTk.PhotoImage(resized_image)
    logo_label = tk.Label(root, image=logo)
    logo_label.image = logo
    logo_label.grid(row=0, column=0, columnspan=5, pady=10)
except Exception as e:
    print(f"Logo load error: {e}")

# ========================== GUI Widgets ==========================
# Barcode
tk.Label(root, text="Barcode:").grid(row=2, column=0)
combo_barcode = ttk.Combobox(root, values=barcodes_list)
combo_barcode.grid(row=2, column=1)
combo_barcode.bind("<<ComboboxSelected>>", autofill_patient_info)

# Patient Name
tk.Label(root, text="Patient Name:").grid(row=2, column=2)
entry_name = tk.Entry(root)
entry_name.grid(row=2, column=3)

# Gender
tk.Label(root, text="Gender:").grid(row=4, column=0)
entry_gender = tk.Entry(root)
entry_gender.grid(row=4, column=1)

# DOB
tk.Label(root, text="DOB:").grid(row=4, column=2)
entry_dob = tk.Entry(root)
entry_dob.grid(row=4, column=3)

# Collected On
tk.Label(root, text="Collected on:").grid(row=6, column=0)
entry_collected_on = tk.Entry(root)
entry_collected_on.grid(row=6, column=1)

# Address
tk.Label(root, text="Address:").grid(row=6, column=2)
entry_address = tk.Entry(root)
entry_address.grid(row=6, column=3)

# Received On
tk.Label(root, text="Recieved on:").grid(row=8, column=0)
entry_recieved_on = tk.Entry(root)
entry_recieved_on.grid(row=8, column=1)

# Hospital Number
tk.Label(root, text="Hospital Number:").grid(row=8, column=2)
entry_hospital_number = tk.Entry(root)
entry_hospital_number.grid(row=8, column=3)

# Reported On
tk.Label(root, text="Reported on:").grid(row=10, column=0)
entry_report_on = tk.Entry(root)
entry_report_on.grid(row=10, column=1)

# Package
tk.Label(root, text="Test Package:").grid(row=10, column=2)
combo_package = ttk.Combobox(root, values=package_names)
combo_package.grid(row=10, column=3)

# ========================== Buttons ==========================
tk.Button(root, text="Generate Report", command=generate_report).grid(row=12, columnspan=5, pady=20)
tk.Button(root, text="Refresh", command=refresh_form).grid(row=14, column=1)
tk.Button(root, text="Exit", command=root.destroy).grid(row=14, column=2)

# Window size
root.geometry("500x400")
root.mainloop()
