"""
Skoda Auto AS Invoice Extractor
Extracts structured line-item data from Skoda Auto AS invoices (PDF).
Outputs a consolidated CSV with INR-style number formatting.
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
    """
    if not value_str or not isinstance(value_str, str):
        return value_str

    value_str = value_str.strip()

    # If the string has both '.' and ',' where ',' is the decimal separator
    # (European format), convert accordingly.
    if '.' in value_str and ',' in value_str:
        # European format: periods are thousands seps, comma is decimal
        # Remove periods (thousands), replace comma with period (decimal)
        converted = value_str.replace('.', '').replace(',', '.')
        # Re-format with commas as thousands separator
        try:
            num = float(converted)
            # Format with commas as thousands and '.' as decimal
            if '.' in converted:
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            return f"{num:,.0f}"
        except ValueError:
            return value_str
    elif ',' in value_str and '.' not in value_str:
        # Only comma present → it's a decimal separator
        converted = value_str.replace(',', '.')
        try:
            num = float(converted)
            decimal_places = len(converted.split('.')[-1])
            return f"{num:,.{decimal_places}f}"
        except ValueError:
            return value_str
    else:
        # Already in standard format or no separators
        return value_str


def eur_str_to_float(value_str: str) -> float:
    """
    Parse a European-formatted number string to a Python float.
    e.g. '2.236,90' → 2236.90, '0,297' → 0.297
    """
    if not value_str or not isinstance(value_str, str):
        return 0.0
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        return float(value_str.replace('.', '').replace(',', '.'))
    elif ',' in value_str:
        return float(value_str.replace(',', '.'))
    else:
        try:
            return float(value_str)
        except ValueError:
            return 0.0


def float_to_inr_str(value: float, decimal_places: int = 2) -> str:
    """Format a float to INR-style string (commas as thousands, period as decimal)."""
    return f"{value:,.{decimal_places}f}"


def eur_qty_to_int(value_str: str) -> int:
    """
    Parse a European-formatted quantity string to a Python int.
    In Skoda invoices, quantities use periods as thousands separators
    and never have decimal parts (always whole pieces).
    e.g. '3.600' → 3600, '60.000' → 60000, '432' → 432
    """
    if not value_str or not isinstance(value_str, str):
        return 0
    # Remove periods (thousands separators) and parse as integer
    cleaned = value_str.strip().replace('.', '')
    try:
        return int(cleaned)
    except ValueError:
        return 0


# ---------- COUNTRY CODE → COUNTRY NAME MAPPING ----------
COUNTRY_MAP: dict[str, str] = {
    "SK": "Slovakia",
    "CZ": "Czech Republic",
    "DE": "Germany",
    "TR": "Turkey",
    "HU": "Hungary",
    "JP": "Japan",
    "PT": "Portugal",
    "RO": "Romania",
    "ES": "Spain",
    "IT": "Italy",
    "FR": "France",
    "PL": "Poland",
    "AT": "Austria",
    "BE": "Belgium",
    "NL": "Netherlands",
    "SE": "Sweden",
    "GB": "United Kingdom",
    "CN": "China",
    "KR": "South Korea",
    "US": "United States",
    "MX": "Mexico",
    "BR": "Brazil",
    "IN": "India",
    "TH": "Thailand",
    "SI": "Slovenia",
    "RS": "Serbia",
    "BA": "Bosnia and Herzegovina",
    "HR": "Croatia",
    "BG": "Bulgaria",
}


# ---------- CORE EXTRACTION LOGIC ----------
def extract_skoda_invoice(pdf_path: str) -> dict:
    """
    Extract all line-item data from a single Skoda Auto AS invoice PDF.
    Returns a dict with header info and a list of line items.
    """
    invoice_number: str = ""
    date_of_supply: str = ""
    currency: str = "EUR"
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        all_text_lines: list[str] = []

        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text_lines.extend(text.split('\n'))

        # --- Extract Header Info ---
        for line in all_text_lines:
            # Invoice Number: "Rechnung Nr. 59638127"
            if not invoice_number:
                inv_match = re.search(r'Rechnung\s+Nr\.\s*(\d+)', line)
                if inv_match:
                    invoice_number = inv_match.group(1)

            # Date of Supply: "Date of supply/Datum der Leistung 15.12.2025"
            if not date_of_supply:
                date_match = re.search(
                    r'Date\s+of\s+supply/Datum\s+der\s+Leistung\s+(\d{2}\.\d{2}\.\d{4})',
                    line
                )
                if date_match:
                    date_of_supply = date_match.group(1)

            # Currency
            if 'EUR' in line and 'Total to be paid' in line:
                currency = "EUR"

        # --- Extract Line Items ---
        # Line items appear in pairs:
        #   Line 1: PartNumber HSCode Quantity UnitPrice TotalPrice
        #   Line 2: CountryCode Description WeightPerPc Divisor /PC
        #
        # Regex for data line 1 (part number line):
        # Part numbers start with alphanumeric, followed by HS code (6 digits),
        # then quantity, unit price, total price
        # Examples:
        #   04C145299B 848350 432 517,80 2.236,90
        #   N 10124501 392690 60.000 16,74 1.004,40
        #   5JN820045B SZB 870829 400 21,83 8.732,00

        # Pattern for the FIRST line of a line item pair
        part_line_pattern = re.compile(
            r'^(.+?)\s+'            # Part number (greedy but lazy to stop before HS code)
            r'(\d{6})\s+'           # HS Code (exactly 6 digits)
            r'([\d.,]+)\s+'         # Quantity
            r'([\d.,]+)\s+'         # Unit Price
            r'([\d.,]+)$'           # Total Price
        )

        # Pattern for the SECOND line of a line item pair
        detail_line_pattern = re.compile(
            r'^([A-Z]{2})\s+'       # Country code (2 uppercase letters)
            r'(.+?)\s+'             # Description (lazy match)
            r'([\d.,]+)\s+'         # Weight per pc
            r'(\d+)\s+/'            # Divisor and /
            r'([A-Za-z]+)$'         # Unit (PC, KG, G, etc.)
        )

        # Footer / non-data markers to skip
        skip_markers = [
            'Strana', 'Page', 'Seite', 'Evnt.', 'IČO:', 'Městský',
            '---', '===', 'Goods', 'Packing costs', 'Total to be paid',
            'Total weight', 'Volume', '###', 'Cust.overview',
            'Poč.obalů', 'Vnitřní modul', 'Name of packaging',
            'PACKING WEIGHT', 'Total weight inc', 'Part number Serial',
            'It is a tax-exempt', 'Es handelt sich', 'Škoda Auto',
            'tř.Václava', 'Mladá Boleslav', '293 01',
            'Daňový doklad', 'Invoice No.', 'Rechnung Nr.', 'Var.symbol',
            'Mat.č./', 'Mat.No./', 'Mat.Nr./',
            'Dodavatel', 'BNP Paribas', 'IBAN', 'BIC/Swift',
            'Č.účtu/', 'Kupující/', 'Příjemce/', 'SKODA AUTO',
            'PRIVATE LIMITED', 'E-1, MIDC', 'VILLAGE NIGOJE',
            'CHAKAN TAL', '410501 PUNE', 'Dod.list/', 'Číslo tran.',
            'Dod.podmínky/', 'Platební', 'Splatnost/', 'Způsob',
            'Místo určení/', 'Datum uskutečnění', 'Date of supply',
            'Dat.vystavení', 'India', 'Partner.spol',
            'Banka/Bank', 'Kód dodávky', 'Obj.č./', 'IČ/Cust',
            'DIČ/Tax', 'No/Bankkonto', 'Deb.', 'FCA',
            'Seatransport', 'India, Pune',
            'A 153', 'A 15S', 'A 0002', 'A 0009', 'A 157', 'A 0000',
        ]

        i = 0
        while i < len(all_text_lines):
            line = all_text_lines[i].strip()

            # Skip blank and footer/header lines
            if not line or any(line.startswith(m) for m in skip_markers):
                i += 1
                continue

            # Try to match a part number line
            match1 = part_line_pattern.match(line)
            if match1 and (i + 1) < len(all_text_lines):
                next_line = all_text_lines[i + 1].strip()
                match2 = detail_line_pattern.match(next_line)

                if match2:
                    # Remove extracting spaces from Part Number (e.g. "05E105561 ROT" -> "05E105561ROT")
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

                    # Parse values
                    quantity = eur_qty_to_int(quantity_str)
                    unit_price = eur_str_to_float(unit_price_str)
                    total_price = eur_str_to_float(total_price_str)
                    weight_per_pc = eur_str_to_float(weight_per_pc_str)
                    divisor = int(divisor_str)

                    # Append suffix to description (MOVED TO SEPARATE COLUMN)
                    full_description = description
                    default_col_text = "(AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION)"

                    # Format numbers as INR-style
                    item = {
                        "Invoice Number": invoice_number,
                        "Invoice Date": date_of_supply,
                        "Mat. NO.": part_number,
                        "Country of Origin": country_code,
                        "HS Code": hs_code,
                        "Description": full_description,
                        "Default": default_col_text,
                        "Quantity": f"{quantity:,}",
                        "Weight per PC (kg)": convert_eur_to_inr_format(weight_per_pc_str),
                        "Unit Price (per batch)": convert_eur_to_inr_format(unit_price_str),
                        "Price per PC": f"{divisor_str} /{unit_str}",
                        "Total Price": convert_eur_to_inr_format(total_price_str),
                        "Currency": currency,
                    }

                    line_items.append(item)
                    i += 2  # Skip both lines
                    continue

            i += 1

    return {
        "invoice_number": invoice_number,
        "date_of_supply": date_of_supply,
        "currency": currency,
        "items": line_items,
    }


def write_csv(output_path: str, all_records: list[dict]) -> None:
    """Write all extracted records to a single CSV file."""
    if not all_records:
        return

    fieldnames = [
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

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in all_records:
            # Wrap Mat. NO. so Excel preserves leading zeros
            safe_record = dict(record)
            mat_no = safe_record.get("Mat. NO.", "")
            if mat_no and mat_no[0] == '0':
                safe_record["Mat. NO."] = f'="{mat_no}"'
            writer.writerow(safe_record)


# ---------- NAGARKOT GUI IMPLEMENTATION ----------
class SkodaExtractorGUI:
    """Skoda Auto AS Invoice Extractor – Nagarkot Branded GUI."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Skoda Auto AS – Invoice Extractor")
        self.root.geometry("1100x700")
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
            text="Skoda Auto AS – Invoice Extractor",
            style="Header.TLabel",
        )
        title_lbl.grid(row=0, column=1, sticky="n")
        subtitle_lbl = ttk.Label(
            header_frame,
            text="Extract line-item data from Skoda Auto AS invoices",
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
            text="Combined (All in one CSV)", 
            variable=self.mode_var, 
            value="combined",
            command=self.toggle_filename_state
        )
        self.rb_combined.pack(side="left", padx=(0, 15))
        
        self.rb_individual = ttk.Radiobutton(
            mode_frame, 
            text="Individual (Separate CSV per invoice)", 
            variable=self.mode_var, 
            value="individual",
            command=self.toggle_filename_state
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
            style="Secondary.TButton"
        )
        self.btn_browse_out.grid(row=1, column=2, sticky="w")

        # --- Output Filename ---
        ttk.Label(output_frame, text="Output Filename:").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_name_var = tk.StringVar(
            value=f"Skoda_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.entry_output_name = ttk.Entry(
            output_frame, textvariable=self.output_name_var, width=50
        )
        self.entry_output_name.grid(row=2, column=1, sticky="ew", padx=(0, 10))
        
        self.lbl_filename_hint = ttk.Label(output_frame, text="(For Combined mode)", foreground="gray")
        self.lbl_filename_hint.grid(row=2, column=2, sticky="w")

        output_frame.columnconfigure(1, weight=1)

        # --- Data Preview ---
        preview_frame = ttk.LabelFrame(
            content_frame,
            text="Data Preview / Processing Queue",
            padding="15",
        )
        preview_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Status", "Items", "Details")
        self.tree = ttk.Treeview(
            preview_frame, columns=cols, show="headings", selectmode="extended"
        )
        self.tree.heading("File Name", text="File Name")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Items", text="Items Found")
        self.tree.heading("Details", text="Details")

        self.tree.column("File Name", width=350, anchor="w")
        self.tree.column("Status", width=100, anchor="center")
        self.tree.column("Items", width=100, anchor="center")
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
            title="Select Skoda Invoice PDFs",
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
                    values=(os.path.basename(fpath), "Pending", "", "")
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
            self.lbl_filename_hint.config(text="(For Combined mode)")

    def clear_files(self) -> None:
        self.selected_files = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.lbl_count.config(text="No files selected")
        self.output_dir_var.set("")
        self.output_name_var.set(
            f"Skoda_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
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

                result = extract_skoda_invoice(fpath)
                items = result["items"]
                count = len(items)
                inv_no = result.get("invoice_number", "N/A")
                inv_date = result.get("date_of_supply", "N/A")
                
                total_items += count

                # --- INDIVIDUAL MODE ---
                if mode == "individual" and items:
                    # Determine filename: InvoiceNumber.csv or OriginalName.csv
                    # Sanitize invoice number for filename
                    safe_inv = "".join(c for c in inv_no if c.isalnum() or c in ('-', '_'))
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
                        values=(fname, "✓ Done", str(count), detail_msg),
                    )

            except Exception as e:
                if row_id:
                    self.tree.item(
                        row_id,
                        values=(fname, "✗ Error", "0", str(e)),
                    )
                self.status_var.set(f"Error processing {fname}: {e}")

            self.root.update_idletasks()

        # Finalize Combined Mode
        if mode == "combined":
            if combined_records:
                out_name = self.output_name_var.get()
                if not out_name:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_name = f"Skoda_Extracted_{timestamp}.csv"
                if not out_name.lower().endswith(".csv"):
                    out_name += ".csv"
                
                output_path = os.path.join(out_dir, out_name)
                try:
                    write_csv(output_path, combined_records)
                    messagebox.showinfo(
                        "Success",
                        f"Combined extraction complete!\n\n"
                        f"Total items: {total_items}\n"
                        f"Saved to: {output_path}"
                    )
                    self.status_var.set(f"Done. Saved to {out_name}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not write combined CSV:\n{e}")
            else:
                self.status_var.set("No data found to combine.")
                if total_items == 0:
                    messagebox.showwarning("No Data", "No items extracted from selected files.")
        
        # Finalize Individual Mode
        else:
             messagebox.showinfo(
                "Success",
                f"Individual extraction complete!\n\n"
                f"Processed {len(self.selected_files)} files.\n"
                f"Total items found: {total_items}\n"
                f"Folder: {out_dir}"
            )
             self.status_var.set(f"Done. Files saved to {out_dir}")

        # Refresh timestamp for next run
        new_ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if self.mode_var.get() == "combined":
             self.output_name_var.set(f"Skoda_Extracted_{new_ts}")

        self._reset_buttons()

    def _reset_buttons(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_select.config(state="normal")

    def run(self) -> None:
        self.root.mainloop()


# ---------- ENTRY POINT ----------
if __name__ == "__main__":
    app = SkodaExtractorGUI()
    app.run()
