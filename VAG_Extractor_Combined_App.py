import os
import sys
import re
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from typing import Optional

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

# ---------- RESOURCE PATH FUNCTION ----------
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

# ---------- EUR → INR NUMBER FORMATTING ----------
def convert_eur_to_inr_format(value_str: str) -> str:
    """
    Convert European number format to Indian/standard format.
    European: 2.236,90 → Standard: 2,236.90
    European: 0,297   → Standard: 0.297
    European: 43.760,64 → Standard: 43,760.64
    If already INR-style (commas as thousands, period as decimal), pass through.
    """
    if not value_str or not isinstance(value_str, str): return value_str
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        last_dot = value_str.rfind('.')
        last_comma = value_str.rfind(',')
        if last_comma > last_dot:
            converted = value_str.replace('.', '').replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            return value_str
    elif ',' in value_str and '.' not in value_str:
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            converted = value_str.replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(parts[1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            return value_str
    else:
        return value_str

def eur_str_to_float(value_str: str) -> float:
    if not value_str or not isinstance(value_str, str): return 0.0
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        if value_str.rfind(',') > value_str.rfind('.'):
            return float(value_str.replace('.', '').replace(',', '.'))
        else:
            return float(value_str.replace(',', ''))
    elif ',' in value_str:
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            return float(value_str.replace(',', '.'))
        else:
            return float(value_str.replace(',', ''))
    else:
        try: return float(value_str)
        except ValueError: return 0.0

def smart_format_number(value_str: str, currency: str) -> str:
    if not value_str or not isinstance(value_str, str): return value_str
    value_str = value_str.strip()
    if currency == "EUR": return convert_eur_to_inr_format(value_str)
    return value_str

def eur_qty_to_int(value_str: str) -> int:
    if not value_str or not isinstance(value_str, str): return 0
    cleaned = value_str.strip().replace('.', '')
    try: return int(cleaned)
    except ValueError: return 0

# ---------- SKODA LOGIC ----------
def extract_skoda_invoice(pdf_path: str) -> dict:
    invoice_number: str = ""
    date_of_supply: str = ""
    currency: str = "EUR"
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        all_text_lines: list[str] = []
        for page in pdf.pages:
            text = page.extract_text()
            if text: all_text_lines.extend(text.split('\n'))

        for line in all_text_lines:
            if not invoice_number:
                inv_match = re.search(r'Rechnung\s+Nr\.\s*(\d+)', line)
                if inv_match: invoice_number = inv_match.group(1)
            if not date_of_supply:
                date_match = re.search(r'Date\s+of\s+supply/Datum\s+der\s+Leistung\s+(\d{2}\.\d{2}\.\d{4})', line)
                if date_match: date_of_supply = date_match.group(1)
            if 'EUR' in line and 'Total to be paid' in line:
                currency = "EUR"

        part_line_pattern = re.compile(r'^(.+?)\s+(\d{6})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$')
        detail_line_pattern = re.compile(r'^([A-Z]{2})\s+(.+?)\s+([\d.,]+)\s+(\d+)\s+/\s*([A-Za-z]+)$')

        skip_markers = ['Strana', 'Page', 'Seite', 'Evnt.', 'IČO:', 'Městský', '---', '===', 'Goods', 'Packing costs', 
                       'Total to be paid', 'Total weight', 'Volume', '###', 'Cust.overview', 'Poč.obalů', 'Vnitřní modul',
                       'Name of packaging', 'PACKING WEIGHT', 'Total weight inc', 'Part number Serial', 'It is a tax-exempt', 
                       'Es handelt sich', 'Škoda Auto', 'tř.Václava', 'Mladá Boleslav', '293 01', 'Daňový doklad', 
                       'Invoice No.', 'Rechnung Nr.', 'Var.symbol', 'Mat.č./', 'Mat.No./', 'Mat.Nr./', 'Dodavatel', 
                       'BNP Paribas', 'IBAN', 'BIC/Swift', 'Č.účtu/', 'Kupující/', 'Příjemce/', 'SKODA AUTO', 
                       'PRIVATE LIMITED', 'E-1, MIDC', 'VILLAGE NIGOJE', 'CHAKAN TAL', '410501 PUNE', 'Dod.list/', 
                       'Číslo tran.', 'Dod.podmínky/', 'Platební', 'Splatnost/', 'Způsob', 'Místo určení/', 
                       'Datum uskutečnění', 'Date of supply', 'Dat.vystavení', 'India', 'Partner.spol', 'Banka/Bank', 
                       'Kód dodávky', 'Obj.č./', 'IČ/Cust', 'DIČ/Tax', 'No/Bankkonto', 'Deb.', 'FCA', 'Seatransport', 
                       'India, Pune', 'A 153', 'A 15S', 'A 0002', 'A 0009', 'A 157', 'A 0000']

        i = 0
        while i < len(all_text_lines):
            line = all_text_lines[i].strip()
            if not line or any(line.startswith(m) for m in skip_markers):
                i += 1
                continue

            match1 = part_line_pattern.match(line)
            if match1 and (i + 1) < len(all_text_lines):
                next_line = all_text_lines[i + 1].strip()
                match2 = detail_line_pattern.match(next_line)

                if match2:
                    part_number = match1.group(1).strip().replace(" ", "")
                    hs_code = match1.group(2).strip()
                    quantity_str = match1.group(3).strip()
                    unit_price_str = match1.group(4).strip()
                    total_price_str = match1.group(5).strip()

                    country_code = match2.group(1).strip()
                    description = match2.group(2).strip()
                    weight_per_pc_str = match2.group(3).strip()
                    divisor_str = match2.group(4).strip()
                    unit_str = match2.group(5).strip()

                    quantity = eur_qty_to_int(quantity_str)
                    
                    item = {
                        "Invoice Number": invoice_number,
                        "Invoice Date": date_of_supply,
                        "Mat. NO.": part_number,
                        "Country of Origin": country_code,
                        "HS Code": hs_code,
                        "Description": description,
                        "Default": "(AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION)",
                        "Quantity": f"{quantity:,}",
                        "Weight per PC (kg)": convert_eur_to_inr_format(weight_per_pc_str),
                        "Unit Price (per batch)": convert_eur_to_inr_format(unit_price_str),
                        "Price per PC": f"{divisor_str} /{unit_str}",
                        "Total Price": convert_eur_to_inr_format(total_price_str),
                        "Currency": currency,
                        "Source": "SKODA AUTO AS"
                    }
                    line_items.append(item)
                    i += 2
                    continue
            i += 1

    return {
        "invoice_number": invoice_number,
        "date_of_supply": date_of_supply,
        "currency": currency,
        "source": "SKODA AUTO AS",
        "items": line_items,
    }

# ---------- VW/AUDI LOGIC ----------
def extract_vw_group_invoice(pdf_path: str, detected_source="UNKNOWN") -> dict:
    invoice_number: str = ""
    date_of_statement: str = ""
    currency: str = ""
    source: str = detected_source
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        for line in first_page_text.split('\n'):
            if not date_of_statement:
                date_match = re.search(r'(\d{2}-\d{2}-\d{4})\s*/\s*\d{4}-\d{2}-\d{2}', line)
                if date_match: date_of_statement = date_match.group(1)

        if "EUR" in first_page_text and "INR" not in first_page_text: currency = "EUR"
        elif "INR" in first_page_text: currency = "INR"
        else: currency = "EUR"

        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue

            for table in tables:
                if not table or len(table) < 2: continue
                header_cell = table[0][0] if table[0] and table[0][0] else ""
                if "PosNr" not in str(header_cell) and "Pos-No" not in str(header_cell):
                    for row in table:
                        if row:
                            for cell in row:
                                if cell and "Invoice" in str(cell) and not invoice_number:
                                    inv_match = re.search(r'Invoice\s*\n(\d+)', str(cell))
                                    if inv_match: invoice_number = inv_match.group(1)
                                if cell and "Date of statement" in str(cell) and not date_of_statement:
                                    ds_match = re.search(r'Date of statement\s*\n(\d{2}-\d{2}-\d{4})', str(cell))
                                    if ds_match: date_of_statement = ds_match.group(1)
                    continue

                i = 1
                while i < len(table):
                    row = table[i]
                    if not row or len(row) < 14:
                        i += 1
                        continue

                    pos_nr = (row[0] or "").strip()
                    if not re.match(r'^\d{6}$', pos_nr):
                        i += 1; continue

                    raw_part_no = (row[1] or "").strip()
                    german_desc = (row[2] or "").strip()
                    country_code = (row[3] or "").strip()
                    hs_code = (row[5] or "").strip()
                    qty_str = (row[8] or "").strip()
                    unit_str = (row[9] or "").strip()
                    net_weight_str = (row[11] or "").strip()
                    price_100_str = (row[12] or "").strip()
                    total_price_str = (row[13] or "").strip()

                    part_number = raw_part_no.replace('\n', '').replace(' ', '')

                    english_desc_parts: list[str] = []
                    j = i + 1
                    while j < len(table):
                        next_row = table[j]
                        if not next_row or len(next_row) < 14:
                            j += 1; continue
                        next_pos = (next_row[0] or "").strip()
                        if re.match(r'^\d{6}$', next_pos): break
                        cell1 = (next_row[1] or "").strip()
                        if "Verpackung" in cell1:
                            j += 1; continue
                        desc_cell = (next_row[2] or "").strip()
                        if desc_cell:
                            clean_desc = desc_cell.replace('\n', ' ').strip()
                            if clean_desc != german_desc.replace('\n', ' ').strip():
                                english_desc_parts.append(clean_desc)
                        j += 1

                    english_desc = " ".join(english_desc_parts) if english_desc_parts else german_desc.replace('\n', ' ')

                    formatted_net_weight = smart_format_number(net_weight_str, currency)
                    formatted_price_100 = smart_format_number(price_100_str, currency)
                    formatted_total = smart_format_number(total_price_str, currency)

                    qty_clean = qty_str.replace(',', '').replace('.', '')
                    try: quantity = int(qty_clean) if qty_clean else 0
                    except ValueError: quantity = 0

                    item = {
                        "Invoice Number": invoice_number,
                        "Invoice Date": date_of_statement,
                        "Part No.": part_number,
                        "Description": english_desc,
                        "Country Code": country_code,
                        "HS Code": hs_code,
                        "Default": "(AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION)",
                        "Quantity": str(quantity),
                        "Unit": unit_str,
                        "Net Weight": formatted_net_weight,
                        "Price/100": formatted_price_100,
                        "Total Price": formatted_total,
                        "Currency": currency,
                        "Source": source,
                    }
                    line_items.append(item)
                    i = j

    return {
        "invoice_number": invoice_number,
        "date_of_statement": date_of_statement,
        "currency": currency,
        "source": source,
        "items": line_items,
    }

# ---------- SINGLE WRITE FUNCTION ----------
def write_csv(output_path: str, all_records: list[dict], source: str) -> None:
    if not all_records:
        return

    # Dynamically select columns based on the source of the data
    if source == "SKODA AUTO AS":
        fieldnames = [
            "Source",
            "Invoice Number",
            "Invoice Date",
            "Country of Origin",
            "HS Code",
            "Mat. NO.",
            "Description",
            "Default",
            "Quantity",
            "Weight per PC (kg)",
            "Unit Price (per batch)",
            "Price per PC",
            "Total Price",
            "Currency",
        ]
    else:
        # VW / Audi structure
        fieldnames = [
            "Source",
            "Invoice Number",
            "Invoice Date",
            "Part No.",
            "Description",
            "Country Code",
            "HS Code",
            "Default",
            "Quantity",
            "Unit",
            "Net Weight",
            "Price/100",
            "Total Price",
            "Currency",
        ]

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        for record in all_records:
            safe_record = dict(record)
            
            part_no = safe_record.get("Part No.", "")
            if part_no and part_no[0] == '0':
                safe_record["Part No."] = f'="{part_no}"'
                
            mat_no = safe_record.get("Mat. NO.", "")
            if mat_no and mat_no[0] == '0':
                safe_record["Mat. NO."] = f'="{mat_no}"'
                
            writer.writerow(safe_record)

# ---------- DETECT SUB-SOURCE (for VW/Audi family) ----------
def _detect_vw_sub_source(pdf_path: str) -> str:
    """Best-effort label for VW-family invoices (Audi AG vs VW AG vs Audi Hungaria)."""
    with pdfplumber.open(pdf_path) as pdf:
        upper = (pdf.pages[0].extract_text() or "").upper()
        if "AUDI HUNGARIA" in upper:
            return "AUDI HUNGARIA"
        if "AUDI AG" in upper:
            return "AUDI AG"
        if "VOLKSWAGEN AG" in upper or "VW AG" in upper:
            return "VOLKSWAGEN AG"
    return "VW/AUDI"

# ---------- GUI IMPLEMENTATION ----------
class CombinedExtractorGUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("VAG Group (VW/Audi/Skoda) – Invoice Extractor")
        self.root.geometry("1100x700")
        self.root.state('zoomed')

        self.bg_color = "#ffffff"
        self.brand_color = "#0056b3"
        self.root.configure(bg=self.bg_color)
        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("TLabel", background=self.bg_color, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Helvetica", 18, "bold"), foreground=self.brand_color, background=self.bg_color)
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 11), foreground="gray", background=self.bg_color)
        self.style.configure("Footer.TLabel", font=("Segoe UI", 9), foreground="#555555", background=self.bg_color)
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=self.brand_color, foreground="white", borderwidth=0)
        self.style.map("Primary.TButton", background=[('active', '#004494')])
        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), background="#f0f0f0", foreground="#333333", borderwidth=1)
        self.style.map("Secondary.TButton", background=[('active', '#e0e0e0')])
        self.style.configure("TLabelframe", background=self.bg_color)
        self.style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.brand_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), foreground=self.brand_color)

        self.setup_ui()
        self.selected_files: list[str] = []

    def setup_ui(self) -> None:
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", pady=20, padx=20)
        header_frame.columnconfigure(0, weight=0); header_frame.columnconfigure(1, weight=1); header_frame.columnconfigure(2, weight=0)

        # Logo
        try:
            if Image and ImageTk:
                logo_path = resource_path("Nagarkot Logo.png")
                if os.path.exists(logo_path):
                    pil_img = Image.open(logo_path)
                    h_pct = 20 / float(pil_img.size[1])
                    w_size = int(float(pil_img.size[0]) * h_pct)
                    pil_img = pil_img.resize((w_size, 20), Image.Resampling.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(pil_img)
                    ttk.Label(header_frame, image=self.logo_img).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
                else: ttk.Label(header_frame, text="[LOGO]", foreground="gray").grid(row=0, column=0, rowspan=2, sticky="w")
        except Exception: pass

        ttk.Label(header_frame, text="VAG Group – Invoice Extractor", style="Header.TLabel").grid(row=0, column=1, sticky="n")
        ttk.Label(header_frame, text="Auto-detects and extracts Line Items for Skoda, VW, Audi AG, Audi Hungaria", style="Subtitle.TLabel").grid(row=1, column=1, sticky="n")

        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", style="Footer.TLabel").pack(side="left", anchor="s")
        self.btn_run = ttk.Button(footer_frame, text="Extract & Generate CSV", command=self.run_extraction, style="Primary.TButton")
        self.btn_run.pack(side="right", padx=10, pady=5)

        content_frame = ttk.Frame(self.root, padding="20 10 20 10")
        content_frame.pack(fill="both", expand=True)

        file_frame = ttk.LabelFrame(content_frame, text="File Selection", padding="15")
        file_frame.pack(fill="x", pady=(0, 15))
        btn_container = ttk.Frame(file_frame); btn_container.pack(fill="x")
        self.btn_select = ttk.Button(btn_container, text="Select PDFs", command=self.select_files, style="Secondary.TButton")
        self.btn_select.pack(side="left", padx=(0, 10))
        ttk.Button(btn_container, text="Clear List", command=self.clear_files, style="Secondary.TButton").pack(side="left")
        self.lbl_count = ttk.Label(btn_container, text="No files selected", style="TLabel")
        self.lbl_count.pack(side="left", padx=(20, 0))

        output_frame = ttk.LabelFrame(content_frame, text="Output Settings", padding="15")
        output_frame.pack(fill="x", pady=(0, 15))

        # --- Invoice Format selector ---
        ttk.Label(output_frame, text="Invoice Format:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        fmt_frame = ttk.Frame(output_frame)
        fmt_frame.grid(row=0, column=1, columnspan=2, sticky="w")
        self.format_var = tk.StringVar(value="vw_audi")
        ttk.Radiobutton(fmt_frame, text="Skoda AS  (Mladá Boleslav / ČSN)", variable=self.format_var, value="skoda_as").pack(side="left", padx=(0, 20))
        ttk.Radiobutton(fmt_frame, text="VW / Audi  (VW AG, Audi AG, Audi Hungaria)", variable=self.format_var, value="vw_audi").pack(side="left")

        # --- Processing Mode selector ---
        ttk.Label(output_frame, text="Processing Mode:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        mode_frame = ttk.Frame(output_frame)
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w")
        self.mode_var = tk.StringVar(value="combined")
        ttk.Radiobutton(mode_frame, text="Combined (All in one superset CSV)", variable=self.mode_var, value="combined", command=self.toggle_filename_state).pack(side="left", padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="Individual (Separate CSV per invoice)", variable=self.mode_var, value="individual", command=self.toggle_filename_state).pack(side="left")

        ttk.Label(output_frame, text="Output Folder:").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        self.output_dir_var = tk.StringVar()
        self.entry_output_dir = ttk.Entry(output_frame, textvariable=self.output_dir_var, width=50)
        self.entry_output_dir.grid(row=2, column=1, sticky="ew", padx=(0, 10))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_dir, style="Secondary.TButton").grid(row=2, column=2, sticky="w")

        ttk.Label(output_frame, text="Output Filename:").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=5)
        self.output_name_var = tk.StringVar(value=f"Combined_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
        self.entry_output_name = ttk.Entry(output_frame, textvariable=self.output_name_var, width=50)
        self.entry_output_name.grid(row=3, column=1, sticky="ew", padx=(0, 10))
        self.lbl_filename_hint = ttk.Label(output_frame, text="(.csv added automatically)", foreground="gray")
        self.lbl_filename_hint.grid(row=3, column=2, sticky="w")
        output_frame.columnconfigure(1, weight=1)

        preview_frame = ttk.LabelFrame(content_frame, text="Data Preview / Processing Queue", padding="15")
        preview_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Detected Source", "Status", "Items", "Details")
        self.tree = ttk.Treeview(preview_frame, columns=cols, show="headings", selectmode="extended")
        for col in cols: self.tree.heading(col, text=col)
        self.tree.column("File Name", width=350, anchor="w")
        self.tree.column("Detected Source", width=150, anchor="center")
        self.tree.column("Status", width=90, anchor="center")
        self.tree.column("Items", width=90, anchor="center")
        self.tree.column("Details", width=300, anchor="w")

        scrollbar_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(content_frame, textvariable=self.status_var, font=("Segoe UI", 9), foreground="#666666", background="#f5f5f5", anchor="w", padding="5 2").pack(fill="x", pady=(10, 0))

    def select_files(self) -> None:
        files = filedialog.askopenfilenames(title="Select VAG Group Invoice PDFs", filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if files:
            self.selected_files = list(files)
            self.lbl_count.config(text=f"{len(self.selected_files)} file(s) selected")
            for row in self.tree.get_children(): self.tree.delete(row)
            for fpath in self.selected_files:
                self.tree.insert("", "end", values=(os.path.basename(fpath), "Pending", "Pending", "", ""))
            self.status_var.set(f"{len(self.selected_files)} file(s) loaded.")
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(self.selected_files[0]))

    def browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder: self.output_dir_var.set(folder)

    def toggle_filename_state(self) -> None:
        if self.mode_var.get() == "individual":
            self.entry_output_name.config(state="disabled")
            self.lbl_filename_hint.config(text="(Auto-named by Invoice No.)")
        else:
            self.entry_output_name.config(state="normal")
            self.lbl_filename_hint.config(text="(.csv added automatically)")

    def clear_files(self) -> None:
        self.selected_files.clear()
        for row in self.tree.get_children(): self.tree.delete(row)
        self.lbl_count.config(text="No files selected")
        self.output_dir_var.set("")
        self.status_var.set("Ready")

    def run_extraction(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select at least one PDF file.")
            return

        out_dir = self.output_dir_var.get() or os.path.dirname(self.selected_files[0])
        mode = self.mode_var.get()
        combined_records: list[dict] = []
        total_items = 0

        self.btn_run.config(state="disabled"); self.btn_select.config(state="disabled")
        self.root.update_idletasks()
        tree_rows = self.tree.get_children()

        for idx, fpath in enumerate(self.selected_files):
            fname = os.path.basename(fpath)
            row_id = tree_rows[idx] if idx < len(tree_rows) else None

            try:
                self.status_var.set(f"Processing: {fname} ...")
                self.root.update_idletasks()

                chosen_format = self.format_var.get()
                if chosen_format == "skoda_as":
                    result = extract_skoda_invoice(fpath)
                else:
                    sub_source = _detect_vw_sub_source(fpath)
                    result = extract_vw_group_invoice(fpath, detected_source=sub_source)
                    result["source"] = sub_source

                items = result["items"]
                count = len(items)
                source = result.get("source", "UNKNOWN")
                inv_no = result.get("invoice_number", "N/A")
                
                total_items += count

                if mode == "individual" and items:
                    safe_inv = "".join(c for c in inv_no if c.isalnum() or c in ('-', '_'))
                    indiv_name = f"{safe_inv}.csv" if safe_inv else f"{os.path.splitext(fname)[0]}_Extracted.csv"
                    # If combined mode receives mixed sources, it will use the superset
                    # But if we want individual mode strictness:
                    write_csv(os.path.join(out_dir, indiv_name), items, source)
                    detail_msg = f"Saved: {indiv_name} ({count} items)"
                else:
                    combined_records.extend(items)
                    detail_msg = f"Invoice: {inv_no} | {count} items"

                if row_id: self.tree.item(row_id, values=(fname, source, "✓ Done", str(count), detail_msg))
            except Exception as e:
                if row_id: self.tree.item(row_id, values=(fname, "—", "✗ Error", "0", str(e)))

            self.root.update_idletasks()

        if mode == "combined" and total_items > 0:
            out_name = self.output_name_var.get().strip() or f"Combined_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            if not out_name.lower().endswith('.csv'): out_name += '.csv'
            out_path = os.path.join(out_dir, out_name)
            try:
                # Determine the primary source for combined mode if completely mixed use superset (we just pass a dummy flag to fall back to a full combined list)
                sources_found = set(r.get("Source", "UNKNOWN") for r in combined_records)
                if len(sources_found) == 1:
                    primary_source = list(sources_found)[0]
                else:
                    primary_source = "MIXED"
                
                if primary_source == "MIXED":
                     # Dynamic superset column extraction
                     csv_cols = list({k: None for r in combined_records for k in r.keys()}.keys())
                     with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                        writer = csv.DictWriter(f, fieldnames=csv_cols, extrasaction="ignore")
                        writer.writeheader()
                        for r in combined_records:
                            safe_record = dict(r)
                            pno = safe_record.get("Part No.", "")
                            if pno and pno[0] == '0': safe_record["Part No."] = f'="{pno}"'
                            mno = safe_record.get("Mat. NO.", "")
                            if mno and mno[0] == '0': safe_record["Mat. NO."] = f'="{mno}"'
                            writer.writerow(safe_record)
                else:
                     write_csv(out_path, combined_records, primary_source)
                messagebox.showinfo("Success", f"Saved {total_items} items to {out_path}")
                self.status_var.set(f"Done. Saved to {out_name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV:\n{e}")
        elif mode == "combined":
            messagebox.showwarning("No Data", "No items extracted.")
        else:
            messagebox.showinfo("Success", f"Individual processing complete. Saved {total_items} items.")
            self.status_var.set(f"Done. Saved to {out_dir}")

        self.btn_run.config(state="normal"); self.btn_select.config(state="normal")
        self.output_name_var.set(f"Combined_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")

    def run(self) -> None:
        self.root.mainloop()

if __name__ == "__main__":
    app = CombinedExtractorGUI()
    app.run()
