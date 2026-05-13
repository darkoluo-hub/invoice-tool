# NDIS Invoice Generator

A Windows desktop tool for generating NDIS tax invoices from an Excel invoice template and an Excel data table.

This project is designed for local workflow use. It lets you select an invoice template, an invoice data file, and an output folder through a simple desktop UI, then generates invoice workbooks automatically. It can also export PDFs and optionally append a participant information PDF to the final invoice PDF.

## Features

- Desktop UI built with Tkinter
- Generate invoices from an Excel template and invoice data workbook
- Group rows by `invoice_no`
- Optional PDF export
- Optional merge of a participant information PDF into the invoice PDF
- Keeps the tax row visible even when tax is zero
- Rebuilds merged description cells to preserve the invoice layout
- Automatically creates a safe output filename if the target file is locked or already exists

## Requirements

- Windows
- Microsoft Excel installed
- Python 3.10 or newer

## Python packages

Install the required packages:

```bash
pip install -r requirements.txt
```

## How to run

You can either:

1. Double-click `NDIS Invoice Generator v1.0.pyw` on Windows, or
2. Run it from the command line:

```bash
pythonw "NDIS Invoice Generator v1.0.pyw"
```

## How it works

1. Select the invoice template `.xlsx` file.
2. Select the invoice data `.xlsx` file.
3. Select the output folder.
4. Optionally tick **Also export PDF**.
5. Optionally tick **Merge Participant Information PDF** and select the PDF.
6. Click **Generate Invoices**.

The tool reads invoice rows from the sheet named `InvoiceData`. If that sheet does not exist, it falls back to the active sheet.

## Expected input structure

The data file should contain headers in row 1.

### Core fields used by the app

These fields are used directly when generating invoices:

- `invoice_no`
- `client_name`
- `client_address`
- `client_city`
- `client_phone`
- `client_email`
- `issue_date`
- `payment_date`
- `ndis_number`
- `type`
- `quantity`
- `unit_price`
- `tax`

### Optional fields

- `output_filename`  
  Custom filename for the generated invoice workbook.

- `full_description` / `line_description` / `b_description`  
  If provided, the exact text is used in the invoice description column.

- `line_detail` / `travel_detail` / `km_detail`  
  Inserts custom detail text after the date and before the category and item reference.

- `line_suffix` / `description_suffix`  
  Appends custom text after the category and item reference.

## Description behaviour

If no full manual description is provided, the app builds the service description automatically in a format similar to:

```text
Worker @ DD/MM/YYYY HH:MM AM - HH:MM AM [Category] [Item]
```

This also supports travel or kilometre-style descriptions such as Google Maps based km lines.

## Important notes

- The current template logic supports up to **5 service rows per invoice**.
- The app uses Excel via `pywin32`, so it is intended for **Windows with desktop Excel installed**.
- PDF export also depends on Excel being available locally.
- Participant information PDF merging uses `pypdf`.

## Typical workflow

A common setup is:

- one Excel invoice template
- one invoice data workbook containing multiple rows
- rows grouped by `invoice_no`
- one output folder for generated `.xlsx` invoices and optional `.pdf` files

## Troubleshooting

### Missing package errors

Install all dependencies again:

```bash
pip install -r requirements.txt
```

### Excel does not open or PDF export fails

Check that:

- Microsoft Excel is installed
- Excel can open the template manually
- the output file is not already open

### No invoices generated

Check that:

- the data file is not empty
- the header row is correct
- the `invoice_no` column contains values

### Too many service rows

The current template supports 5 service lines only. If one invoice contains more than 5 rows, the script will stop and show an error.

## Repository notes

This project is currently focused on a practical local invoicing workflow rather than packaging or cross-platform deployment.

Possible future improvements:

- sample input workbook
- sample invoice template
- packaged `.exe` release
- better validation for required columns
- configurable template row count
- logging to file

## Disclaimer

Please test generated invoices carefully before using them in production. Always review invoice content, support item descriptions, tax handling, and participant details before sending invoices to clients, plan managers, or platform providers.
