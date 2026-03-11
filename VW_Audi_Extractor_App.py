"""
VW / Audi AG / Audi Hungaria – Invoice Extractor
Extracts structured line-item data from Volkswagen-group invoices (PDF).
Supports: Volkswagen AG, Audi AG, Audi Hungaria Zrt.
Outputs CSV with INR-style number formatting.
"""

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
    if not value_str or not isinstance(value_str, str):
        return value_str

    value_str = value_str.strip()

    # Detect EUR format: both '.' and ',' present, where comma comes AFTER the
    # last period → European style (e.g. "164.675,59")
    if '.' in value_str and ',' in value_str:
        last_dot = value_str.rfind('.')
        last_comma = value_str.rfind(',')

        if last_comma > last_dot:
            # European format: periods are thousands, comma is decimal
            converted = value_str.replace('.', '').replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            # Already INR format (comma is thousands, period is decimal)
            return value_str

    elif ',' in value_str and '.' not in value_str:
        # Ambiguous: could be EUR decimal or INR thousands.
        # Heuristic: if digits after the comma <= 3 AND there is only one comma
        # AND the part after comma is 1-3 digits, treat as EUR decimal.
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            # Likely European decimal (e.g. "0,297" or "517,80")
            converted = value_str.replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(parts[1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            # Likely already INR thousands (e.g. "2,300")
            return value_str
    else:
        return value_str


def eur_str_to_float(value_str: str) -> float:
    """
    Parse a European-formatted number string to a Python float.
    e.g. '2.236,90' → 2236.90, '0,297' → 0.297
    Also handles INR-style strings.
    """
    if not value_str or not isinstance(value_str, str):
        return 0.0
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        last_dot = value_str.rfind('.')
        last_comma = value_str.rfind(',')
        if last_comma > last_dot:
            # European: remove dots, replace comma with dot
            return float(value_str.replace('.', '').replace(',', '.'))
        else:
            # INR: remove commas
            return float(value_str.replace(',', ''))
    elif ',' in value_str:
        # Could be EUR decimal or INR thousands
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            return float(value_str.replace(',', '.'))
        else:
            return float(value_str.replace(',', ''))
    else:
        try:
            return float(value_str)
        except ValueError:
            return 0.0


def float_to_inr_str(value: float, decimal_places: int = 2) -> str:
    """Format a float to INR-style string (commas as thousands, period as decimal)."""
    return f"{value:,.{decimal_places}f}"


def smart_format_number(value_str: str, currency: str) -> str:
    """
    Intelligently format a number string based on the invoice currency.
    EUR invoices use European number formatting (dots=thousands, comma=decimal).
    INR invoices already use standard formatting (commas=thousands, period=decimal).
    Always output INR-style.
    """
    if not value_str or not isinstance(value_str, str):
        return value_str
    value_str = value_str.strip()
    if currency == "EUR":
        return convert_eur_to_inr_format(value_str)
    else:
        # INR already – just pass through (already in correct format)
        return value_str


# ---------- DETECT INVOICE SOURCE ----------
def detect_source(first_page_text: str) -> str:
    """Detect whether the invoice is from VW, Audi AG, or Audi Hungaria."""
    upper = first_page_text.upper()
    if "AUDI HUNGARIA" in upper:
        return "AUDI HUNGARIA"
    elif "AUDI AG" in upper:
        return "AUDI AG"
    elif "VOLKSWAGEN AG" in upper:
        return "VOLKSWAGEN AG"
    else:
        return "UNKNOWN"


# ---------- CORE TABLE-BASED EXTRACTION LOGIC ----------
def extract_vw_group_invoice(pdf_path: str) -> dict:
    """
    Extract all line-item data from a VW-group invoice PDF using table extraction.
    Works for: Volkswagen AG, Audi AG, Audi Hungaria Zrt.
    Returns a dict with header info and a list of line items.
    """
    invoice_number: str = ""
    date_of_statement: str = ""
    currency: str = ""
    source: str = ""
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        # --- Detect source from first page ---
        first_page_text = pdf.pages[0].extract_text() or ""
        source = detect_source(first_page_text)

        # --- Extract header from first page text ---
        for line in first_page_text.split('\n'):
            # Invoice number
            if not invoice_number:
                # Header table row 0 contains "Rechnung\nInvoice\n<NUMBER>"
                # Also appears as standalone number on page 1
                # Try the simple pattern from text first
                pass

            # Date of statement: "15-12-2025 / 2025-12-15" on page 1 overview
            if not date_of_statement:
                date_match = re.search(
                    r'(\d{2}-\d{2}-\d{4})\s*/\s*\d{4}-\d{2}-\d{2}',
                    line
                )
                if date_match:
                    date_of_statement = date_match.group(1)

        # --- Extract header from first page TABLE 1 ---
        first_page_tables = pdf.pages[0].extract_tables()
        if first_page_tables:
            # The first table on page 1 is the summary/condition table
            pass

        # --- Detect currency from page 1 text ---
        if "EUR" in first_page_text and "INR" not in first_page_text:
            currency = "EUR"
        elif "INR" in first_page_text:
            currency = "INR"
        else:
            currency = "EUR"  # Default to EUR

        # --- Process each data page (skip page 1 = summary) ---
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                continue

            for table in tables:
                if not table or len(table) < 2:
                    continue

                # Check if this is a line-item table by looking at header row
                header_cell = table[0][0] if table[0] and table[0][0] else ""
                if "PosNr" not in str(header_cell) and "Pos-No" not in str(header_cell):
                    # Could be the header table – extract invoice info
                    for row in table:
                        if row:
                            for cell in row:
                                if cell and "Invoice" in str(cell) and not invoice_number:
                                    inv_match = re.search(
                                        r'Invoice\s*\n(\d+)', str(cell)
                                    )
                                    if inv_match:
                                        invoice_number = inv_match.group(1)
                                if cell and "Date of statement" in str(cell) and not date_of_statement:
                                    ds_match = re.search(
                                        r'Date of statement\s*\n(\d{2}-\d{2}-\d{4})',
                                        str(cell)
                                    )
                                    if ds_match:
                                        date_of_statement = ds_match.group(1)
                    continue

                # This IS a line-item table
                # Column indices (0-based) from inspection:
                # 0: PosNr    1: Part-No    2: Description (German)
                # 3: Country  4: (empty)    5: Duty-code
                # 6: Order-No 7: HU-No      8: Qty
                # 9: Unit    10: BZA       11: Net weight
                # 12: Price/100            13: Total price

                # Process rows AFTER the header (R0)
                i = 1
                while i < len(table):
                    row = table[i]
                    if not row or len(row) < 14:
                        i += 1
                        continue

                    pos_nr = (row[0] or "").strip()

                    # A data row starts with a 6-digit position number like "000010"
                    if not re.match(r'^\d{6}$', pos_nr):
                        i += 1
                        continue

                    # --- Main data row ---
                    raw_part_no = (row[1] or "").strip()
                    german_desc = (row[2] or "").strip()
                    country_code = (row[3] or "").strip()
                    hs_code = (row[5] or "").strip()
                    qty_str = (row[8] or "").strip()
                    unit_str = (row[9] or "").strip()
                    net_weight_str = (row[11] or "").strip()
                    price_100_str = (row[12] or "").strip()
                    total_price_str = (row[13] or "").strip()

                    # Clean up part number: join ALL newline segments,
                    # then remove ALL spaces.
                    # Part numbers can have suffixes on newlines, e.g.:
                    #   "4M0 863 879 H\nHCK" → "4M0863879HHCK"
                    #   "571 601 025 AB\nFZZ" → "571601025ABFZZ"
                    part_number = raw_part_no.replace('\n', '').replace(' ', '')

                    # --- Collect English description from subsequent rows ---
                    english_desc_parts: list[str] = []
                    j = i + 1
                    while j < len(table):
                        next_row = table[j]
                        if not next_row or len(next_row) < 14:
                            j += 1
                            continue

                        next_pos = (next_row[0] or "").strip()

                        # If we hit another position number, stop
                        if re.match(r'^\d{6}$', next_pos):
                            break

                        # Skip "Verpackung Programmliefg." packing cost rows
                        cell1 = (next_row[1] or "").strip()
                        if "Verpackung" in cell1:
                            j += 1
                            continue

                        # The English description is in column 2 of
                        # subsequent rows (where pos is None/empty)
                        desc_cell = (next_row[2] or "").strip()
                        if desc_cell:
                            # Clean up multi-line descriptions
                            clean_desc = desc_cell.replace('\n', ' ').strip()
                            # Skip if it's identical to the German description
                            if clean_desc != german_desc.replace('\n', ' ').strip():
                                english_desc_parts.append(clean_desc)

                        j += 1

                    english_desc = " ".join(english_desc_parts) if english_desc_parts else german_desc.replace('\n', ' ')

                    # --- Format numbers ---
                    formatted_net_weight = smart_format_number(net_weight_str, currency)
                    formatted_price_100 = smart_format_number(price_100_str, currency)
                    formatted_total = smart_format_number(total_price_str, currency)

                    # Quantity: may have commas as thousands (INR) e.g. "3,000" or "20,000"
                    # or periods as thousands (EUR) e.g. "3.600"
                    qty_clean = qty_str.replace(',', '').replace('.', '')
                    try:
                        quantity = int(qty_clean) if qty_clean else 0
                    except ValueError:
                        quantity = 0

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
                    i = j  # Jump past the English description rows

    return {
        "invoice_number": invoice_number,
        "date_of_statement": date_of_statement,
        "currency": currency,
        "source": source,
        "items": line_items,
    }


# ---------- CSV OUTPUT ----------
def write_csv(output_path: str, all_records: list[dict]) -> None:
    """Write all extracted records to a single CSV file."""
    if not all_records:
        return

    fieldnames = [
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
        "Source",
    ]

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in all_records:
            # Wrap Part No. so Excel preserves leading zeros
            safe_record = dict(record)
            part_no = safe_record.get("Part No.", "")
            if part_no and part_no[0] == '0':
                safe_record["Part No."] = f'="{part_no}"'
            writer.writerow(safe_record)


# ---------- NAGARKOT GUI IMPLEMENTATION ----------
class VWAudiExtractorGUI:
    """VW / Audi Invoice Extractor – Nagarkot Branded GUI."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("VW / Audi – Invoice Extractor")
        self.root.geometry("1200x750")
        self.root.state('zoomed')

        # Nagarkot brand palette
        self.bg_color = "#ffffff"
        self.brand_color = "#0056b3"
        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style()
        self.style.theme_use('clam')

        # --- Style configuration ---
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure(
            "TLabel", background=self.bg_color, font=("Segoe UI", 10)
        )
        self.style.configure(
            "Header.TLabel",
            font=("Helvetica", 18, "bold"),
            foreground=self.brand_color,
            background=self.bg_color,
        )
        self.style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 11),
            foreground="gray",
            background=self.bg_color,
        )
        self.style.configure(
            "Footer.TLabel",
            font=("Segoe UI", 9),
            foreground="#555555",
            background=self.bg_color,
        )
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background=self.brand_color,
            foreground="white",
            borderwidth=0,
            focuscolor=self.brand_color,
        )
        self.style.map("Primary.TButton", background=[('active', '#004494')])
        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            background="#f0f0f0",
            foreground="#333333",
            borderwidth=1,
        )
        self.style.map("Secondary.TButton", background=[('active', '#e0e0e0')])
        self.style.configure("TLabelframe", background=self.bg_color)
        self.style.configure(
            "TLabelframe.Label",
            background=self.bg_color,
            foreground=self.brand_color,
            font=("Segoe UI", 10, "bold"),
        )
        self.style.configure(
            "Treeview", font=("Segoe UI", 9), rowheight=25
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            foreground=self.brand_color,
        )

        self.setup_ui()
        self.selected_files: list[str] = []

    # ----- UI SETUP -----
    def setup_ui(self) -> None:
        # ---------- HEADER ----------
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", pady=20, padx=20)
        header_frame.columnconfigure(0, weight=0)
        header_frame.columnconfigure(1, weight=1)
        header_frame.columnconfigure(2, weight=0)

        # Logo (Left)
        try:
            if Image and ImageTk:
                logo_path = resource_path("Nagarkot Logo.png")
                if os.path.exists(logo_path):
                    pil_img = Image.open(logo_path)
                    h_pct = 20 / float(pil_img.size[1])
                    w_size = int(float(pil_img.size[0]) * h_pct)
                    pil_img = pil_img.resize(
                        (w_size, 20), Image.Resampling.LANCZOS
                    )
                    self.logo_img = ImageTk.PhotoImage(pil_img)
                    logo_lbl = ttk.Label(header_frame, image=self.logo_img)
                    logo_lbl.grid(
                        row=0, column=0, rowspan=2, sticky="w", padx=(0, 20)
                    )
                else:
                    print("Warning: Nagarkot Logo.png not found.")
                    ttk.Label(
                        header_frame, text="[LOGO]", foreground="gray"
                    ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
            else:
                ttk.Label(
                    header_frame, text="[PIL Missing]", foreground="red"
                ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
        except Exception as e:
            print(f"Error loading logo: {e}")
            ttk.Label(
                header_frame, text="[LOGO ERROR]", foreground="red"
            ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))

        # Title (Center)
        title_lbl = ttk.Label(
            header_frame,
            text="VW / Audi – Invoice Extractor",
            style="Header.TLabel",
        )
        title_lbl.grid(row=0, column=1, sticky="n")
        subtitle_lbl = ttk.Label(
            header_frame,
            text="Extract line-item data from Volkswagen AG · Audi AG · Audi Hungaria invoices",
            style="Subtitle.TLabel",
        )
        subtitle_lbl.grid(row=1, column=1, sticky="n")

        # ---------- FOOTER (Packed first to reserve bottom space) ----------
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")

        copyright_lbl = ttk.Label(
            footer_frame,
            text="© Nagarkot Forwarders Pvt Ltd",
            style="Footer.TLabel",
        )
        copyright_lbl.pack(side="left", anchor="s")

        self.btn_run = ttk.Button(
            footer_frame,
            text="Extract & Generate CSV",
            command=self.run_extraction,
            style="Primary.TButton",
        )
        self.btn_run.pack(side="right", padx=10, pady=5)

        # ---------- MAIN CONTENT ----------
        content_frame = ttk.Frame(self.root, padding="20 10 20 10")
        content_frame.pack(fill="both", expand=True)

        # --- File Selection ---
        file_frame = ttk.LabelFrame(
            content_frame, text="File Selection", padding="15"
        )
        file_frame.pack(fill="x", pady=(0, 15))

        btn_container = ttk.Frame(file_frame)
        btn_container.pack(fill="x")

        self.btn_select = ttk.Button(
            btn_container,
            text="Select PDFs",
            command=self.select_files,
            style="Secondary.TButton",
        )
        self.btn_select.pack(side="left", padx=(0, 10))

        self.btn_clear = ttk.Button(
            btn_container,
            text="Clear List",
            command=self.clear_files,
            style="Secondary.TButton",
        )
        self.btn_clear.pack(side="left")

        self.lbl_count = ttk.Label(
            btn_container, text="No files selected", style="TLabel"
        )
        self.lbl_count.pack(side="left", padx=(20, 0))

        # --- Output Settings ---
        output_frame = ttk.LabelFrame(
            content_frame, text="Output Settings", padding="15"
        )
        output_frame.pack(fill="x", pady=(0, 15))

        # --- Processing Mode (Combined vs Individual) ---
        ttk.Label(output_frame, text="Processing Mode:").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=5
        )

        mode_frame = ttk.Frame(output_frame)
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="w")

        self.mode_var = tk.StringVar(value="combined")

        self.rb_combined = ttk.Radiobutton(
            mode_frame,
            text="Combined (All in one combined Csv)",
            variable=self.mode_var,
            value="combined",
            command=self.toggle_filename_state,
        )
        self.rb_combined.pack(side="left", padx=(0, 15))

        self.rb_individual = ttk.Radiobutton(
            mode_frame,
            text="Individual (Separate Csv per invoice)",
            variable=self.mode_var,
            value="individual",
            command=self.toggle_filename_state,
        )
        self.rb_individual.pack(side="left")

        # --- Output Folder ---
        ttk.Label(output_frame, text="Output Folder:").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_dir_var = tk.StringVar()
        self.entry_output_dir = ttk.Entry(
            output_frame, textvariable=self.output_dir_var, width=50
        )
        self.entry_output_dir.grid(row=1, column=1, sticky="ew", padx=(0, 10))

        self.btn_browse_out = ttk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_dir,
            style="Secondary.TButton",
        )
        self.btn_browse_out.grid(row=1, column=2, sticky="w")

        # --- Output Filename ---
        ttk.Label(output_frame, text="Output Filename:").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_name_var = tk.StringVar(
            value=f"VW_Audi_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.entry_output_name = ttk.Entry(
            output_frame, textvariable=self.output_name_var, width=50
        )
        self.entry_output_name.grid(row=2, column=1, sticky="ew", padx=(0, 10))

        self.lbl_filename_hint = ttk.Label(
            output_frame, text="(.csv added automatically)", foreground="gray"
        )
        self.lbl_filename_hint.grid(row=2, column=2, sticky="w")

        output_frame.columnconfigure(1, weight=1)

        # --- Data Preview ---
        preview_frame = ttk.LabelFrame(
            content_frame,
            text="Data Preview / Processing Queue",
            padding="15",
        )
        preview_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Source", "Status", "Items", "Details")
        self.tree = ttk.Treeview(
            preview_frame, columns=cols, show="headings", selectmode="extended"
        )
        self.tree.heading("File Name", text="File Name")
        self.tree.heading("Source", text="Source")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Items", text="Items Found")
        self.tree.heading("Details", text="Details")

        self.tree.column("File Name", width=350, anchor="w")
        self.tree.column("Source", width=130, anchor="center")
        self.tree.column("Status", width=90, anchor="center")
        self.tree.column("Items", width=90, anchor="center")
        self.tree.column("Details", width=400, anchor="w")

        scrollbar_y = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            content_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            foreground="#666666",
            background="#f5f5f5",
            anchor="w",
            padding="5 2",
        )
        status_bar.pack(fill="x", pady=(10, 0))

    # ----- FILE SELECTION -----
    def select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select VW / Audi Invoice PDFs",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if files:
            self.selected_files = list(files)
            self.lbl_count.config(
                text=f"{len(self.selected_files)} file(s) selected"
            )
            # Clear and populate treeview
            for row in self.tree.get_children():
                self.tree.delete(row)
            for fpath in self.selected_files:
                self.tree.insert(
                    "", "end",
                    values=(os.path.basename(fpath), "—", "Pending", "", ""),
                )
            self.status_var.set(
                f"{len(self.selected_files)} file(s) loaded. "
                "Click 'Extract & Generate CSV' to process."
            )
            # Auto-set output folder if empty
            if not self.output_dir_var.get():
                first_dir = os.path.dirname(self.selected_files[0])
                self.output_dir_var.set(first_dir)

    def browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def toggle_filename_state(self) -> None:
        """Enable/Disable filename entry based on mode."""
        if self.mode_var.get() == "individual":
            self.entry_output_name.config(state="disabled")
            self.lbl_filename_hint.config(text="(Auto-named by Invoice No.)")
        else:
            self.entry_output_name.config(state="normal")
            self.lbl_filename_hint.config(text="(.csv added automatically)")

    def clear_files(self) -> None:
        self.selected_files = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.lbl_count.config(text="No files selected")
        self.output_dir_var.set("")
        self.status_var.set("File list cleared.")

    # ----- RUN EXTRACTION -----
    def run_extraction(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(
                "No Files", "Please select at least one PDF file."
            )
            return

        # Output setup
        out_dir = self.output_dir_var.get()
        if not out_dir:
            out_dir = os.path.dirname(self.selected_files[0])
            self.output_dir_var.set(out_dir)

        mode = self.mode_var.get()
        combined_records: list[dict] = []
        total_items = 0

        self.btn_run.config(state="disabled")
        self.btn_select.config(state="disabled")
        self.root.update_idletasks()

        tree_rows = self.tree.get_children()

        for idx, fpath in enumerate(self.selected_files):
            fname = os.path.basename(fpath)
            row_id = tree_rows[idx] if idx < len(tree_rows) else None

            try:
                self.status_var.set(f"Processing: {fname} ...")
                self.root.update_idletasks()

                result = extract_vw_group_invoice(fpath)
                items = result["items"]
                count = len(items)
                inv_no = result.get("invoice_number", "N/A")
                source = result.get("source", "UNKNOWN")

                total_items += count

                # --- INDIVIDUAL MODE ---
                if mode == "individual" and items:
                    safe_inv = "".join(
                        c for c in inv_no if c.isalnum() or c in ('-', '_')
                    )
                    if safe_inv:
                        indiv_name = f"{safe_inv}.csv"
                    else:
                        base = os.path.splitext(fname)[0]
                        indiv_name = f"{base}_Extracted.csv"

                    indiv_path = os.path.join(out_dir, indiv_name)
                    write_csv(indiv_path, items)
                    detail_msg = f"Saved: {indiv_name} ({count} items)"

                # --- COMBINED MODE ---
                else:
                    combined_records.extend(items)
                    detail_msg = f"Invoice: {inv_no} | {count} items"

                if row_id:
                    self.tree.item(
                        row_id,
                        values=(
                            fname, source, "✓ Done", str(count), detail_msg
                        ),
                    )

            except Exception as e:
                if row_id:
                    self.tree.item(
                        row_id,
                        values=(fname, "—", "✗ Error", "0", str(e)),
                    )
                self.status_var.set(f"Error processing {fname}: {e}")

            self.root.update_idletasks()

        # Finalize Combined Mode
        if mode == "combined":
            if combined_records:
                out_name = self.output_name_var.get().strip()
                if not out_name:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_name = f"VW_Audi_Extracted_{timestamp}"
                # Always ensure .csv extension
                if out_name.lower().endswith(".csv"):
                    out_name = out_name[:-4]  # Strip it, we re-add below
                out_name += ".csv"

                output_path = os.path.join(out_dir, out_name)
                try:
                    write_csv(output_path, combined_records)
                    messagebox.showinfo(
                        "Success",
                        f"Combined extraction complete!\n\n"
                        f"Total items: {total_items}\n"
                        f"Saved to: {output_path}",
                    )
                    self.status_var.set(f"Done. Saved to {out_name}")
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Could not write combined CSV:\n{e}"
                    )
            else:
                self.status_var.set("No data found to combine.")
                if total_items == 0:
                    messagebox.showwarning(
                        "No Data", "No items extracted from selected files."
                    )

        # Finalize Individual Mode
        else:
            messagebox.showinfo(
                "Success",
                f"Individual extraction complete!\n\n"
                f"Processed {len(self.selected_files)} files.\n"
                f"Total items found: {total_items}\n"
                f"Folder: {out_dir}",
            )
            self.status_var.set(f"Done. Files saved to {out_dir}")

        # Refresh timestamp for next run
        new_ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if self.mode_var.get() == "combined":
            self.output_name_var.set(f"VW_Audi_Extracted_{new_ts}")

        self._reset_buttons()

    def _reset_buttons(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_select.config(state="normal")

    def run(self) -> None:
        self.root.mainloop()


# ---------- ENTRY POINT ----------
if __name__ == "__main__":
    app = VWAudiExtractorGUI()
    app.run()
