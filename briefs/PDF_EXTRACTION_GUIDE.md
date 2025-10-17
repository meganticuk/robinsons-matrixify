# PDF Stock Inventory Extraction Script

## Overview
This guide provides a Python script to extract stock inventory data from Happy Socks PDF reports (T2T Mercury system) and convert it to CSV format.

## PDF Structure Analysis

The PDF contains a Master Stock Listing with the following columns:
- **T2TRef**: Product reference number
- **Style Name**: Product name
- **Style Description**: Product code/description
- **Splr**: Supplier (HAP - Happy Socks)
- **Supp Ref.**: Supplier reference code
- **Dept**: Department (ACC - Accessories)
- **Type**: Product type (GBO - Gift Box, SOC - Socks)
- **Colour**: Color variant
- **Sizes And Quantities**: S, M, L, XL columns with quantities
- **Totals**: Total quantity across all sizes

### Data Challenges
1. **Multi-row entries**: Each product has size quantities on separate rows (S/M/L/XL)
2. **Page headers**: Headers repeat on each page
3. **Summary rows**: Subtotals and grand totals need filtering
4. **Type sections**: Data grouped by product type

## Installation Requirements

```bash
pip install tabula-py pandas openpyxl
```

**Note**: tabula-py requires Java to be installed on your system.

Alternative (if Java not available):
```bash
pip install pdfplumber pandas openpyxl
```

## Python Script - Option 1: Using tabula-py (Recommended)

```python
import tabula
import pandas as pd
import os

def extract_inventory_from_pdf(pdf_path, output_csv="inventory_output.csv"):
    """
    Extract inventory data from Happy Socks PDF stock listing.

    Args:
        pdf_path (str): Path to the PDF file
        output_csv (str): Output CSV filename
    """
    print(f"Reading PDF: {pdf_path}")

    # Extract all tables from PDF
    tables = tabula.read_pdf(
        pdf_path,
        pages='all',
        multiple_tables=True,
        pandas_options={'header': None}
    )

    # Combine all tables
    all_data = []
    for table in tables:
        all_data.append(table)

    df = pd.concat(all_data, ignore_index=True)

    # Clean the data
    df = clean_inventory_data(df)

    # Save to CSV
    df.to_csv(output_csv, index=False)
    print(f"Data extracted successfully to: {output_csv}")

    return df

def clean_inventory_data(df):
    """
    Clean and structure the extracted data.
    """
    # Remove completely empty rows
    df = df.dropna(how='all')

    # Remove rows that are headers (contain "T2TRef")
    df = df[~df.iloc[:, 0].astype(str).str.contains('T2TRef', na=False)]

    # Remove summary rows (contain "Totals")
    df = df[~df.iloc[:, 0].astype(str).str.contains('Totals', na=False)]
    df = df[~df.iloc[:, 0].astype(str).str.contains('Type ', na=False)]

    # Reset index
    df = df.reset_index(drop=True)

    return df

def main():
    # Set the PDF filename
    pdf_file = "HAPPYSOCKS-16-10-25-0915.pdf"

    # Check if file exists
    if not os.path.exists(pdf_file):
        print(f"Error: {pdf_file} not found in current directory")
        return

    # Extract data
    df = extract_inventory_from_pdf(pdf_file)

    # Display summary
    print(f"\nExtraction Summary:")
    print(f"Total rows extracted: {len(df)}")
    print(f"\nFirst few rows:")
    print(df.head())

if __name__ == "__main__":
    main()
```

## Python Script - Option 2: Using pdfplumber (More Control)

```python
import pdfplumber
import pandas as pd
import re

def extract_inventory_with_pdfplumber(pdf_path, output_csv="inventory_output.csv"):
    """
    Extract inventory data using pdfplumber for better control.

    Args:
        pdf_path (str): Path to the PDF file
        output_csv (str): Output CSV filename
    """
    print(f"Reading PDF: {pdf_path}")

    all_rows = []
    headers = ['T2TRef', 'Style Name', 'Style Description', 'Splr', 'Supp Ref.',
               'Dept', 'Type', 'Colour', 'S', 'M', 'L', 'XL', 'Totals']

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"Processing page {page_num}...")

            # Extract tables from page
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    # Skip header rows
                    if row and row[0] and 'T2TRef' in str(row[0]):
                        continue

                    # Skip summary rows
                    if row and row[0] and ('Totals' in str(row[0]) or 'Type ' in str(row[0])):
                        continue

                    # Skip empty rows
                    if not row or all(cell is None or str(cell).strip() == '' for cell in row):
                        continue

                    all_rows.append(row)

    # Create DataFrame
    df = pd.DataFrame(all_rows)

    # Set proper column names
    if len(df.columns) >= len(headers):
        df.columns = headers + [f'Extra_{i}' for i in range(len(df.columns) - len(headers))]

    # Save to CSV
    df.to_csv(output_csv, index=False)
    print(f"Data extracted successfully to: {output_csv}")

    return df

def merge_size_rows(df):
    """
    Advanced: Merge size rows into single product rows.
    Each product has multiple rows for different sizes (S/M/L/XL).
    """
    merged_data = []
    current_product = None

    for idx, row in df.iterrows():
        if pd.notna(row['T2TRef']) and str(row['T2TRef']).strip():
            # New product entry
            if current_product:
                merged_data.append(current_product)

            current_product = {
                'T2TRef': row['T2TRef'],
                'Style Name': row['Style Name'],
                'Style Description': row['Style Description'],
                'Supplier': row['Splr'],
                'Supplier Ref': row['Supp Ref.'],
                'Department': row['Dept'],
                'Type': row['Type'],
                'Colour': row['Colour'],
                'Size_S': row['S'] if pd.notna(row['S']) else 0,
                'Size_M': row['M'] if pd.notna(row['M']) else 0,
                'Size_L': row['L'] if pd.notna(row['L']) else 0,
                'Size_XL': row['XL'] if pd.notna(row['XL']) else 0,
                'Total': row['Totals'] if pd.notna(row['Totals']) else 0
            }
        elif current_product:
            # Size continuation row
            if pd.notna(row['S']):
                current_product['Size_S'] = row['S']
            if pd.notna(row['M']):
                current_product['Size_M'] = row['M']
            if pd.notna(row['L']):
                current_product['Size_L'] = row['L']
            if pd.notna(row['XL']):
                current_product['Size_XL'] = row['XL']

    # Add last product
    if current_product:
        merged_data.append(current_product)

    return pd.DataFrame(merged_data)

def main():
    pdf_file = "HAPPYSOCKS-16-10-25-0915.pdf"

    # Extract data
    df = extract_inventory_with_pdfplumber(pdf_file, "inventory_raw.csv")

    print(f"\nRaw extraction complete: {len(df)} rows")
    print(f"\nFirst few rows:")
    print(df.head(20))

    # Optional: Merge size rows
    print("\nMerging size rows into single product entries...")
    merged_df = merge_size_rows(df)
    merged_df.to_csv("inventory_merged.csv", index=False)
    print(f"Merged data saved: {len(merged_df)} products")

if __name__ == "__main__":
    main()
```

## Usage

1. **Place the script in the same folder as your PDF**
2. **Run the script**:
   ```bash
   python extract_inventory.py
   ```

3. **Output files**:
   - `inventory_output.csv` - Raw extracted data
   - `inventory_merged.csv` - Products with sizes merged into single rows

## Expected Output Structure

### Raw CSV Format
```csv
T2TRef,Style Name,Style Description,Splr,Supp Ref.,Dept,Type,Colour,S,M,L,XL,Totals
001189,HOLIDAY 3P BOX,XMAS08-4001,HAP,XMAS08-4001,ACC,GBO,MULTI,,,3,,3
001692,HAPPY BIRTHDAY,XBDA08-2700,HAP,XBDA08-2700,ACC,GBO,BLUE,,,1,,1
```

### Merged CSV Format
```csv
T2TRef,Style Name,Style Description,Supplier,Supplier Ref,Department,Type,Colour,Size_S,Size_M,Size_L,Size_XL,Total
001189,HOLIDAY 3P BOX,XMAS08-4001,HAP,XMAS08-4001,ACC,GBO,MULTI,0,0,3,0,3
001692,HAPPY BIRTHDAY,XBDA08-2700,HAP,XBDA08-2700,ACC,GBO,BLUE,0,0,1,0,1
```

## Data Summary from Current PDF

- **Total Items**: 1,610 units
- **Product Types**:
  - GBO (Gift Boxes): 111 units
  - SOC (Socks): 1,499 units
- **Branches**: 0, A, B (consolidated)
- **Supplier**: HAP (Happy Socks)
- **Date**: 16/10/2025 @ 09:12

## Troubleshooting

### Java Not Installed (for tabula-py)
If you get a Java error:
1. Install Java: https://www.java.com/download/
2. Or use pdfplumber option (no Java required)

### Incorrect Column Detection
- Adjust `pandas_options` in tabula.read_pdf()
- Use `lattice=True` for bordered tables
- Use `stream=True` for non-bordered tables

### Missing Data
- Check if PDF is text-based (not scanned image)
- Verify PDF opens correctly in a PDF reader
- Try adjusting extraction area with `area` parameter

## Advanced Options

### Extract Specific Pages Only
```python
tables = tabula.read_pdf(pdf_path, pages='1-5')
```

### Export to Excel
```python
df.to_excel("inventory_output.xlsx", index=False)
```

### Filter by Product Type
```python
gift_boxes = df[df['Type'] == 'GBO']
socks = df[df['Type'] == 'SOC']
```

## Next Steps

1. Run the script to extract data
2. Review the CSV output
3. Import into your inventory management system
4. Set up automated extraction for future reports

## Notes

- The script handles multi-page PDFs automatically
- Headers and summary rows are filtered out
- Size quantities are preserved across columns
- All product metadata is captured (T2TRef, Style, Color, etc.)
