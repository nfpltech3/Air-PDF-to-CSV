# Integrated DO + Invoice to CSV Converter User Guide

## Introduction
The Integrated DO + Invoice to CSV Converter is an automated tool designed for Nagarkot Forwarders Pvt Ltd. to extract text data from multiple logistics Vendor invoices (such as MIAL, Air India, Dachser, DHL, Anil Mantra, etc.) and map them securely to corresponding internal Job Numbers via a Job Register. It outputs a standardized, ready-to-import CSV file containing all merged data.

## How to Use

### 1. Launching the App
1. Locate the `Air_PDF_to_CSV.exe` application file (found in your `dist/` folder if freshly built).
2. Double-click the `.exe` file to launch the application in full-screen mode.

### 2. The Workflow (Step-by-Step)
1. **Select Job Register**: Click the `Select Job Register` button. Choose your Job Register file (supported: `.csv`, `.xls`, `.xlsx`).
   - *Note: Ensure your Job Register contains columns like `BE No`, `BOE No`, `Job No`, `HAWB/HBL No`, and `AWB/BL No.`. Dates in the file will automatically be handled.*
2. **Select PDFs (Invoice/DO)**: Click the `Select PDFs (Invoice/DO)` button. You can select one or multiple PDF files at once.
   - *Note: The PDFs must contain selectable text (not purely scanned images) for the extraction engine to work correctly.*
3. **Process & Generate CSV**: Click the `▶ Process & Generate CSV` button. 
   - *Note: You must have both the Job Register and PDFs selected before processing.*
4. **Retrieve Output**: Check your success prompt. The merged CSV file is saved automatically into a folder named `CSV_Output` in the same directory as the executable.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Job Register** | Loads the mapping file used to find Job Numbers. | `.csv`, `.xls`, `.xlsx` |
| **Select PDFs (Invoice/DO)** | Selects the vendor invoices or delivery orders for extraction. | `.pdf` (Text-based) |
| **▶ Process & Generate CSV** | Starts the text extraction and matching process. | N/A |
| **Processing Log Area** | Displays real-time updates on which PDF is being processed, and logs mappings or failures for debugging. | Text Output |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| `No Job Register file selected.` | You did not select a mapping file. | Click "Select Job Register" and pick a valid CSV/Excel file. |
| `No PDFs selected.` | You did not select any invoices to process. | Click "Select PDFs" and highlight the files you wish to process. |
| `Unsupported file type selected.` | The selected Job Register is not a supported format. | Ensure the file is exactly `.csv`, `.xls`, or `.xlsx`. |
| `Error loading job register: [error]` | The file structure is corrupt or being locked by another program. | Close the Job Register if it is open in Excel, and try again. |
| `Failed to process [FileName]` | The PDF could not be read. | Ensure the PDF is not password-protected and contains recognizable text (not a scanned image). |
| `No match found` (In Log/Output) | The system could not map the HAWB/MAWB or BE No to a Job No. | Verify that the reference number exists in your Job Register and does not contain abnormal characters. |
| `No valid data extracted from PDFs.` | None of the provided PDFs matched known vendor formats (MIAL, Air India, or standard DO). | Ensure the PDFs are standard logistics invoices that the system supports. |
