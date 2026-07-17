# Importer Invoice Mapping & Extraction Guide

This document outlines how the General Item Mismatch Checker parses, maps, and processes invoice files (Excel/PDF) for different importers.

---

## 1. General Importers (Dynamic Excel Mapping) - e.g., BBraun, etc.
For any importer not explicitly listed with custom logic below (such as **BBraun**), the application uses **dynamic Excel header mapping** configured inside the master Google Sheet.

*   **Source File**: Excel (`.xlsx`, `.xls`)
*   **Configuration**: Sourced from the `Invoice_Header_Mappings` tab in the Master Google Sheet.
    *   **Importer**: Name of the importer (e.g., `BBraun`).
    *   **Source_Header**: The exact column header text in the invoice Excel.
    *   **Target_Field**: The target field in the checklist template (e.g., `Inv No`, `Date`, `Product Desc`, `Quantity`, `Unit`, `Currency`, `Rate`, `Country of origin`).
*   **Behavior**:
    *   Reads the Excel sheet.
    *   Maps columns matching `Source_Header` (case-insensitive) to `Target_Field`.
    *   **Model Selection Prompt**: If the `Model` column cannot be found or is not mapped, a popup window will request the user to select the Model column manually from a dropdown of available headers.
    *   **Date formatting**: Gracefully converts date objects to `DD-MM-YYYY` string format.

---

## 2. ADVICS (Custom Excel Mapping)
ADVICS has a fixed Excel layout mapping that does not rely on the `Invoice_Header_Mappings` configuration.

*   **Source File**: Excel (`.xlsx`, `.xls`)
*   **Header Metadata Extraction**:
    *   Scans the top 30 rows of the sheet for exact text indicators:
        *   `INVOICE NO. :` $\rightarrow$ Maps to `Inv No` (`col_1`)
        *   `DATE :` $\rightarrow$ Maps to `Date` (`col_2`)
        *   `Country of origin :` $\rightarrow$ Maps to `Country of origin` (`col_32`)
*   **Line Items Table Extraction**:
    *   Locates the start of the table by finding the row containing both `PART NO.` and `DESCRIPTION`.
    *   Extracts line items from the following columns:
        *   `PART NO.` $\rightarrow$ `Model`
        *   `DESCRIPTION` $\rightarrow$ `Product Desc`
        *   `Qty` $\rightarrow$ `Quantity`
        *   `Unit Price` (any column containing "Unit Price") $\rightarrow$ `Rate` (If `(THB)` is present in header, default currency is set to `THB`)
        *   **Unit (UOM)**: Sourced from the column directly following `Qty`. If the value is `"pcs"` or `"pcs."`, it is normalized to `"NOS"`.

---

## 3. ARJO (Custom Excel Mapping)
ARJO uses a fixed Excel layout mapping, very similar to ADVICS but with slightly different header names.

*   **Source File**: Excel (`.xlsx`, `.xls`)
*   **Header Metadata Extraction**:
    *   Scans the top rows for specific embedded text:
        *   A cell containing `"Invoice No."` $\rightarrow$ Extracts the number (e.g., `"1006256741"`) and maps to `Inv No` (`col_1`).
        *   A cell containing `"Invoice date :"` $\rightarrow$ Extracts the date text (e.g., `"27 MAY 2026"`) and formats it to `DD-MM-YYYY` string format. Maps to `Date` (`col_2`).
*   **Line Items Table Extraction**:
    *   Locates the start of the table by finding the row containing both `"Part No"` and `"Qty"`.
    *   Extracts line items from the following columns:
        *   `Part No` $\rightarrow$ `Model`
        *   `Qty` $\rightarrow$ `Quantity`
        *   `Price` $\rightarrow$ `Rate`
        *   `Country Of Origin` $\rightarrow$ `Country of Origin` (`col_32`)
        *   **Currency**: Dynamically extracted from the `"Price"` header by parsing the text inside the parentheses (e.g., `"Price (GBP)"` yields `"GBP"`).
        *   **Unit**: Hardcoded as `"NOS"` for all lines.
    *   **Note**: Other fields like `Product Desc` and `CTH` are naturally loaded from the master sheet based on the extracted `Model` code.

---

## 4. ANSELL (Custom PDF Mapping)
ANSELL invoices are commercial invoices in PDF format. They are parsed using `pdfplumber` to extract tables and text.

*   **Source File**: PDF (`.pdf` only)
*   **Header Metadata Extraction**:
    *   Finds the first table containing the text `"Commercial Invoice"` in the first cell $\rightarrow$ Extracts `Invoice No` (row 3, col 1) and `Invoice Date` (row 3, col 2).
*   **Line Items Table Extraction**:
    *   Identifies item tables containing `"Product Description"` in their header row.
    *   Extracts values from the columns matching the following keywords:
        *   `Product Code` (or a cell containing `"Product"` and `"Code"`) $\rightarrow$ `Model`
        *   `Product Description` $\rightarrow$ Raw Product Description (used for construction)
        *   `Country` $\rightarrow$ `Country of Origin`
        *   `HTS/HS` $\rightarrow$ `HS_Code`
        *   `Case` $\rightarrow$ `Ship Case (Carton)`
        *   `Qty` (from subheader) $\rightarrow$ `Quantity`
        *   `UOM` (from subheader) $\rightarrow$ `Unit`
        *   `Price` (from subheader) $\rightarrow$ `Unit Price` (used for construction)
        *   `Cur` $\rightarrow$ `Currency`
        *   `Value` $\rightarrow$ Invoice line value (used for rate calculation)
*   **Calculations & Formats**:
    *   **Rate**: Calculated dynamically as $\text{Rate} = \text{Value} \div \text{Qty}$ (rounded to 4 decimal places). Strips commas from numerical values before division.
    *   **Product Desc [Auto]**: Dynamically constructed at checklist generation using the template:
        $$\text{PRODUCT CODE. } \{Model\}\ \{raw\ description\}\ (\{Generic\ Description\})\ \text{[QTY } \{Ship\ Case\}\ \text{CTN @ } \{Currency\}\ \{Unit\ Price\}\ \text{PER CTN]}$$
        *   Standardizes spacing to exactly one space after `"PRODUCT CODE."`.
        *   Strips `.00` from the Ship Case Quantity if present (e.g., `79.00` becomes `79`).
        *   If the model is not found in the master sheet, it uses the same template but leaves the generic description empty: `()`.
        *   When submitting a new model from the pending approval queue, the `Generic Description` entered is automatically merged into the `Product Desc` template before sending to the Google Sheet.
    *   **CTH (HS Code)**: Sourced from the master sheet. If the model is not found or has no CTH in the master sheet, falls back to the `HTS/HS` extracted from the invoice.

---

## 5. AVIAT (Custom PDF & Excel Extraction)
AVIAT supports both PDF invoices and extracted Excel sheets. Both bypass the manual mapping configuration dialog.

### Option A: PDF Invoice Upload
*   **Source File**: PDF (`.pdf` only)
*   **Logic**: Uses the custom `aviat_inv_extractor.py` module.
*   **Metadata & Header Extraction**:
    *   **Invoice No**: Scans the first page for the pattern `CI-XXXX-XX`.
    *   **Date**: Scans the first page for the date pattern (e.g. `DD-MMM-YYYY`), converted to standard `DD-MM-YYYY`.
    *   **Currency**: Dynamically detected by finding the word `"CURRENCY"` on page 1, and reading the last token of the line immediately below it (typically `"USD"`).
*   **Line Items Table Extraction**:
    *   **Model**: Extracted from the item list as `Part No`.
    *   **Product Desc**: Extracted from the `Description` block.
    *   **Quantity**: Extracted from the quantity field.
    *   **Rate**: Extracted from the `Unit Price` field.
    *   **Country of Origin**: Scans subsequent lines for `CoO : [Country]` (e.g. `CoO : CN`).
    *   **CTH (HS Code)**: Scans subsequent lines for `CT : [HS Code]` (e.g. `CT : 8471.49.00`), strips trailing decimals.
    *   **Unit**: Hardcoded as `"NOS"`.

### Option B: Extracted Excel Upload
*   **Source File**: Excel (`.xlsx`, `.xls`)
*   **Column Mappings**:
    *   `Commercial Invoice No` $\rightarrow$ `Inv No` (`col_1`)
    *   `Document Date (Long)` $\rightarrow$ `Date` (`col_2`), parsed to standard `DD-MM-YYYY`
    *   `Part No` $\rightarrow$ `Model` (`col_24`)
    *   `Quantity` $\rightarrow$ `Quantity` (`col_5`)
    *   `Unit Price` $\rightarrow$ `Rate` (`col_8`)
    *   `CoO` $\rightarrow$ `Country of Origin` (`col_32`)
    *   `Currency` $\rightarrow$ `Currency` (`col_7`)
    *   **Unit**: Hardcoded as `"NOS"` (`col_6`)
*   **Validation Constraints**:
    *   **CoO Check**: If the `CoO` column is missing, or if any part row contains a blank CoO, the system halts with a popup instructing the user to add/fill the CoO details.
    *   **Currency Check**: If the `Currency` column is missing, or if any part row contains a blank Currency, the system halts with a popup instructing the user to add/fill the Currency details.

