"""
CYS Multi-PDF Requirement Extractor (Optimized with Enhanced GUI)

Features:
---------
✔ Parse multiple PDF files and extract all CYS GUIDs (e.g., CYS-HSM, CYS-SHE, etc.)
✔ Support legacy lines with multiple GUIDs like "Legacy GUID: CYS-HSM_abc / CYS-SHE_xyz"
✔ Extract requirement details, ignoring headers/footers and tables
✔ Dynamically extract and append Cadence number (e.g., Cadence 22.22.142)
✔ Highlight differences in details for same GUID across cadence versions
✔ Append new cadence columns to previously saved Excel files
✔ GUI: Add/Remove PDFs, update existing Excel, display status, show progress bar
✔ Auto save Excel file in script's folder with timestamped name
✔ Automatically open the generated Excel file (for new file creation only)
"""

import fitz  # PyMuPDF
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, END, ttk
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
import difflib
import subprocess
import threading
import time

# Patterns
table_pattern = re.compile(r'^\|.*\|$')
header_footer_pattern = re.compile(r"GM Confidential|Page \d+|^\s*\d+\s*$", re.IGNORECASE)
heading_guid_pattern = re.compile(r"^\d+(\.\d+)*\s+.*GUID:", re.IGNORECASE)
any_guid_pattern = re.compile(r".*GUID:\s*CYS-[A-Z]+[_-]?[\w-]+", re.IGNORECASE)
guid_extract_pattern = re.compile(r"CYS-[A-Z]+[_-]?[\w-]+", re.IGNORECASE)

def is_file_locked(filepath):
    if not os.path.exists(filepath):
        return False
    try:
        os.rename(filepath, filepath)
        return False
    except OSError:
        return True

def format_paragraph(lines):
    return ' '.join(line.strip() for line in lines if line.strip())

def extract_cadence(text):
    match = re.search(r"\b(\d+\.\d+\.\d+)\b", text)
    return f"Cadence {match.group(1)}" if match else "Cadence Unknown"

def extract_requirements_final(pdf_path):
    doc = fitz.open(pdf_path)
    cadence = extract_cadence(doc[0].get_text())
    requirements = []
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        lines = [line.strip() for line in page.get_text().split('\n') if line.strip()]
        i = 0
        while i < len(lines):
            line = lines[i]
            guid_matches = guid_extract_pattern.findall(line)
            if guid_matches:
                details = []
                j = i + 1
                while j < len(lines):
                    next_line = lines[j].strip()
                    if guid_extract_pattern.search(next_line):
                        break
                    if not any_guid_pattern.match(next_line) and not header_footer_pattern.search(next_line) and not table_pattern.match(next_line) and not heading_guid_pattern.match(next_line):
                        details.append(next_line)
                    j += 1
                if details:
                    detail = format_paragraph(details)
                    info_type = "Information" if "(information only)" in line.lower() else "Requirement"
                    for guid in guid_matches:
                        requirements.append([guid.strip(), info_type, detail, cadence, ""])
                i = j
            else:
                i += 1
    return requirements, cadence

def highlight_differences(old_text, new_text):
    diff = list(difflib.ndiff(old_text.split(), new_text.split()))
    highlighted = ' '.join([f'*{w[2:]}*' if w.startswith('+ ') else w[2:] for w in diff if not w.startswith('- ')])
    return highlighted

def save_to_excel(requirements_all, output_path, update_existing=False):
    if update_existing and os.path.exists(output_path):
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        existing_ids = {ws.cell(row=i, column=1).value: i for i in range(2, ws.max_row + 1)}
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        headers = ["Requirement ID", "Requirement/Information"]
        ws.append(headers)
        existing_ids = {}

    for req_id, info_type, detail, cadence, service in requirements_all:
        if cadence not in headers:
            headers.insert(-1 if "HSE Service" in headers else len(headers), cadence)
            ws.cell(row=1, column=headers.index(cadence) + 1).value = cadence
        if "HSE Service" not in headers:
            headers.append("HSE Service")
            ws.cell(row=1, column=headers.index("HSE Service") + 1).value = "HSE Service"
        if req_id in existing_ids:
            row = existing_ids[req_id]
            col = headers.index(cadence) + 1
            existing_detail = ws.cell(row=row, column=col).value or ""
            if existing_detail != detail:
                ws.cell(row=row, column=col).value = highlight_differences(existing_detail, detail)
        else:
            row = ws.max_row + 1
            ws.cell(row=row, column=1).value = req_id
            ws.cell(row=row, column=2).value = info_type
            for head in headers[2:]:
                col_idx = headers.index(head) + 1
                if head == cadence:
                    ws.cell(row=row, column=col_idx).value = detail
                elif head == "HSE Service":
                    ws.cell(row=row, column=col_idx).value = service
            existing_ids[req_id] = row

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_length + 5, 50)

    for _ in range(5):
        if is_file_locked(output_path):
            time.sleep(1)
        else:
            break
    else:
        messagebox.showerror("File Locked", f"Unable to update {output_path} because it is open or locked.")
        return

    wb.save(output_path)

    if not update_existing:
        subprocess.run(["start", output_path], shell=True)

def generate_excel_filename(first_pdf_path):
    with open(first_pdf_path, 'rb') as f:
        content = f.read().decode('latin1', errors='ignore')
        match = re.search(r'CYS-[A-Z]+', content)
        prefix = match.group(0).split('-')[1].lower() if match else "guid"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"cys-{prefix}_{timestamp}.xlsx"

def browse_pdfs():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    for file in files:
        if file not in file_listbox.get(0, END):
            file_listbox.insert(END, file)

def remove_selected():
    for i in file_listbox.curselection()[::-1]:
        file_listbox.delete(i)

def extract_all():
    def task():
        files = file_listbox.get(0, END)
        if not files:
            messagebox.showwarning("No Files", "Please add PDF files.")
            return
        all_requirements = []
        progress_var.set(0)
        progress_bar['maximum'] = len(files)
        for idx, pdf_file in enumerate(files):
            extracted, _ = extract_requirements_final(pdf_file)
            all_requirements.extend(extracted)
            progress_var.set(idx + 1)
        if not all_requirements:
            messagebox.showinfo("No Data", "No valid requirements found.")
            return
        output_dir = os.path.dirname(files[0])
        output_file = generate_excel_filename(files[0])
        output_path = os.path.join(output_dir, output_file)
        if update_var.get():
            existing_file = filedialog.askopenfilename(title="Select Existing Excel File", filetypes=[("Excel Files", "*.xlsx")])
            if not existing_file:
                return
            output_path = existing_file
            save_to_excel(all_requirements, output_path, update_existing=True)
            status_label.config(text="✅ Excel updated.")
        else:
            save_to_excel(all_requirements, output_path, update_existing=False)
            status_label.config(text="✅ New Excel created.")
        progress_var.set(0)
    threading.Thread(target=task).start()

# GUI Setup
root = tk.Tk()
root.title("CYS Multi-PDF Requirement Extractor")
root.geometry("700x550")
root.configure(bg='#f0f4f8')

style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#0078d4", foreground="#333", font=('Segoe UI', 10))
style.configure("TLabel", background="#f0f4f8", font=('Segoe UI', 10))
style.configure("TCheckbutton", background="#f0f4f8", font=('Segoe UI', 10))

frame = ttk.Frame(root)
frame.pack(padx=20, pady=20, fill='both', expand=True)

file_listbox = Listbox(frame, selectmode=tk.MULTIPLE, width=80, height=10, font=('Consolas', 9))
file_listbox.pack(pady=5)

btn_frame = ttk.Frame(frame)
btn_frame.pack(pady=5)

add_btn = ttk.Button(btn_frame, text="Add PDF Files", command=browse_pdfs)
add_btn.grid(row=0, column=0, padx=5)
remove_btn = ttk.Button(btn_frame, text="Remove Selected", command=remove_selected)
remove_btn.grid(row=0, column=1, padx=5)

update_var = tk.BooleanVar()
update_check = ttk.Checkbutton(frame, text="Update Existing Excel", variable=update_var)
update_check.pack(pady=5)

extract_btn = ttk.Button(frame, text="Extract Requirements", command=extract_all)
extract_btn.pack(pady=10)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(frame, variable=progress_var, orient='horizontal', length=500, mode='determinate')
progress_bar.pack(pady=5)

status_label = ttk.Label(frame, text="")
status_label.pack(pady=5)

root.mainloop()
