# VAG Group (VW/Audi/Skoda) – Invoice Extractor User Guide

## Introduction
The **VAG Group Invoice Extractor** is a standalone desktop application designed to automatically extract line-item data from PDF invoices and convert it into a structured CSV (Excel) format. It is explicitly built for the logistics and freight forwarding industry to process commercial invoices from Skoda Auto AS, Volkswagen AG, Audi AG, and Audi Hungaria. 

This tool eliminates manual data entry by extracting Part Numbers, Descriptions, HS Codes, Quantities, Weights, and Pricing directly from the PDF tables.

**Who is it for?**
Customs clearance teams, freight forwarders, and logistics coordinators processing VAG Group automotive parts shipments.

---

## Getting Started

### Prerequisites
* **No installation required.** This is a standalone Windows executable (`.exe`).
* **No internet connection required.** All processing is done locally on your machine.
* **Input Files:** Only standard layout PDF invoices from the supported VAG group entities are accepted.

### Launching the App
1. Download the `VAG_Extractor_Combined.exe` file.
2. Double-click the `.exe` file to launch the application.
3. The application will open in a full-screen, branded window.

---

## Step-by-Step Usage Guide

### 1. Select Your Invoice Format
Before loading files, you must tell the application what type of invoices you are processing. 
* Look at the top left of the application under **Output Settings -> Invoice Format**.
* Select **Skoda AS (Mladá Boleslav / ČSN)** if your invoices are text-based ČSN invoices from Skoda Auto AS.
* Select **VW / Audi (VW AG, Audi AG, Audi Hungaria)** if your invoices are table-based invoices from Volkswagen, Audi AG, or Audi Hungaria.
  * *Note: Do not mix different formats in a single batch. Process Skoda invoices in one batch, and VW/Audi invoices in another.*

### 2. Choose Processing Mode
* **Combined (All in one superset CSV):** Merges all line items from all selected PDFs into a single, master CSV file. Best for bulk uploading into customs software.
* **Individual (Separate CSV per invoice):** Creates a distinct CSV file for every PDF you selected. Best for keeping records strictly separated by invoice.

### 3. Load Your Invoices
1. Click the **Select PDFs** button in the "File Selection" section.
2. Browse your computer and select one or multiple `.pdf` invoice files.
3. Click **Open**. The files will appear in the "Data Preview / Processing Queue" list at the bottom of the screen with a status of "Pending".

### 4. Configure Output Location
1. Under "Output Settings", check the **Output Folder** path. By default, it saves the CSV in the same folder where your PDFs are located.
2. If you want to save them elsewhere, click **Browse...** and select a new destination folder.
3. If using "Combined" mode, you can type a custom name in the **Output Filename** box (e.g., `Jan_2026_Shipment`).

### 5. Run the Extraction
1. Click the blue **Extract & Generate CSV** button in the bottom right corner.
2. The application will process each file. Watch the "Status" column in the queue to see it change from *Pending* to *✓ Done*.
3. A success popup will appear telling you exactly how many total items were extracted.
4. Open the generated CSV file in Excel to review your data.

---

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Invoice Format (Radio Buttons)** | Tells the engine which PDF parser to use. Crucial for accurate data extraction depending on the entity. | Select exact match for the source PDF. |
| **Processing Mode (Radio Buttons)** | Determines if output should be 1 Master CSV or Multiple Individual CSVs. | Selection |
| **Select PDFs (Button)** | Opens the file browser to load invoices. | `.pdf` files only. |
| **Clear List (Button)** | Removes all currently loaded PDFs from the queue so you can start fresh. | N/A |
| **Output Folder (Text / Browse)** | The directory where the final CSV(s) will be saved. | Valid Windows File Path |
| **Output Filename (Text Input)** | The name of the file (Only available in *Combined* mode). | Text (no need to type `.csv` at the end) |
| **Extract & Generate CSV (Button)** | Starts the PDF reading and data extraction sequence. | N/A |
| **Processing Queue (Data Table)** | Shows the real-time status, detected sender, and row count extracted for every file in the batch. | Visual Reference |

---

## Troubleshooting & Validations

If you see an error or experience unexpected behavior, check this table:

| Message / Behavior | What it means | Solution |
| :--- | :--- | :--- |
| **"No Files: Please select at least one PDF file."** | You clicked Extract without loading any PDFs. | Click **Select PDFs** and choose your files first. |
| **"No Data: No items extracted."** | The application scanned the PDFs but found 0 matching line items. | 1. Ensure you selected the correct **Invoice Format** radio button for these specific PDFs. <br> 2. Ensure the PDF is an actual commercial invoice, not a Bill of Lading or Customs document. |
| **Status: "✗ Error" in the queue** | The PDF is corrupted, heavily encrypted, or fundamentally broken. | Open the PDF manually in a viewer to ensure it is readable and not a scanned image (must be digital text). |
| **Missing Columns in Excel** | You opened the CSV but data looks squished into one column. | This is an Excel setting, not an app error. Ensure Excel is set to use commas (`,`) as list separators, or use Excel's "Text to Columns" feature. |
| **Part Numbers missing leading zeros (e.g., `00010` becomes `10`)** | Excel automatically deletes leading zeros. | The application actually writes `="00010"` to force Excel to keep it. If viewing raw CSV, you will see `="[Part]"`. This is intentional to protect data integrity. |
