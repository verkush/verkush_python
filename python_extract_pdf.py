import fitz  # PyMuPDF
import pandas as pd
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

def extract_valid_requirements_only(pdf_path):
    doc = fitz.open(pdf_path)
    requirements = []

    current_id = None
    current_details = []
    cr_number = ""
    info_type = "Requirement"

    for page in doc:
        lines = page.get_text().split('\n')

        # Remove headers and footers
        lines = [
            line for line in lines
            if not re.search(r"GM Confidential|Page \d+|^\s*\d+\s*$", line)
        ]

        for i, line in enumerate(lines):
            line = line.strip()
            if "GUID: CYS-" in line:
                # Check next few lines for early skip (avoid headings or GUID chains)
                next_lines = lines[i+1:i+4]
                has_next_guid = any("GUID: CYS-" in l for l in next_lines)
                has_substantial_text = any(len(l.strip()) > 20 for l in next_lines)

                if has_next_guid or not has_substantial_text:
                    # Skip heading-like GUID
                    current_id, current_details, cr_number = None, [], ""
                    continue

                # Store previous if exists
                if current_id and current_details:
                    full_id = f"{current_id} / {cr_number}" if cr_number else current_id
                    details_text = "\n".join(current_details).strip()
                    requirements.append([full_id, details_text, info_type, ""])

                # New ID
                match = re.search(r"(CYS-[\w\-]+)", line)
                current_id = match.group(1) if match else None
                cr_match = re.search(r"CR\s+\d+", line)
                cr_number = cr_match.group(0) if cr_match else ""
                info_type = "Information" if "(information only)" in line.lower() else "Requirement"
                current_details = []
            elif current_id:
                if "CR " in line and not cr_number:
                    cr_match = re.search(r"CR\s+\d+", line)
                    cr_number = cr_match.group(0) if cr_match else ""
                elif not re.match(r'^\|.*\|$', line):  # Skip table-like rows
                    current_details.append(line.strip())

    # Final save
    if current_id and current_details:
        full_id = f"{current_id} / {cr_number}" if cr_number else current_id
        details_text = "\n".join(current_details).strip()
        requirements.append([full_id, details_text, info_type, ""])

    return requirements

def save_to_excel(data, pdf_path):
    df = pd.DataFrame(data, columns=["Requirement ID", "Details", "Requirement/Information", "HSE Service"])
    output_dir = Path("Requirements_Extracted")
    output_dir.mkdir(exist_ok=True)
    file_name = Path(pdf_path).stem + ".xlsx"
    output_path = output_dir / file_name

    df.to_excel(output_path, index=False)

    # Format the Excel file
    wb = load_workbook(output_path)
    ws = wb.active

    # Bold headers
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    # Wrap text and auto-adjust column width
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = min(max_length + 2, 80)  # cap width
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(output_path)
    return output_path

def process_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        return

    try:
        extracted = extract_valid_requirements_only(pdf_path)
        output_path = save_to_excel(extracted, pdf_path)
        messagebox.showinfo("Success", f"Excel saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI Setup
root = tk.Tk()
root.title("Requirement Extractor")
root.geometry("400x150")

frame = tk.Frame(root, pady=20)
frame.pack()

label = tk.Label(frame, text="Select a PDF file to extract requirements:")
label.pack(pady=10)

btn = tk.Button(frame, text="Select PDF", command=process_pdf)
btn.pack()

root.mainloop()
