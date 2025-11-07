import pandas as pd
import xlwings as xw
import random
import os

def generate_dummy_data(num_rows=100):
    headers = ['Catoragory', 'Groceraries', 'Brand', 'MoQ', 'Unit Price', 'Qty', 'Metric', 'Total']
    categories = ['Staples', 'Snacks', 'Beverages', 'Spices', 'Cleaning']
    groceries = ['Rice', 'Dal', 'Soap', 'Tea', 'Biscuits', 'Salt', 'Sugar', 'Detergent']
    brands = ['BrandA', 'BrandB', 'BrandC', 'BrandD', 'BrandE']
    metrics = ['kg', 'litre', 'pcs']

    data = []
    for _ in range(num_rows):
        cat = random.choice(categories)
        item = random.choice(groceries)
        brand = random.choice(brands)
        moq = random.randint(1, 10)
        price = round(random.uniform(10, 100), 2)
        qty = random.randint(1, 20)
        metric = random.choice(metrics)
        total = round(price * qty, 2)
        shortfall = moq - qty if qty < moq else ''
        data.append([cat, item, brand, moq, price, qty, metric, total, shortfall])

    df = pd.DataFrame(data, columns=headers + ['Shortfall'])
    return df

def write_or_open_excel(filename="Provision_List.xlsm"):
    # Open existing workbook or create + populate if missing.
    if os.path.exists(filename):
        print(f"Workbook '{filename}' exists. Opening for changes...")
        wb = xw.Book(filename)
        # try to select the named sheet if present
        try:
            sheet = wb.sheets['Provision_List']
        except Exception:
            sheet = wb.sheets[0]
    else:
        print(f"Workbook '{filename}' not found. Creating and populating...")
        df = generate_dummy_data()
        wb = xw.Book()
        sheet = wb.sheets[0]
        sheet.name = "Provision_List"
        sheet.range("A1").options(index=False).value = df
        wb.save(filename)

    # Determine number of rows from the sheet's used region (includes header)
    try:
        region = sheet.range("A1").current_region
        total_rows = int(region.shape[0])
    except Exception:
        # fallback: if df exists (we just created it), use its length + header
        total_rows = (len(df) + 1) if 'df' in locals() else 1

    # Detect if the sheet includes a leftmost index column (happens when df was written with index=True)
    try:
        header_vals = sheet.range("A1").expand('right').value
    except Exception:
        header_vals = None

    # header_vals is usually a list of header strings; check where our first expected header sits
    start_col_offset = 0
    if isinstance(header_vals, list):
        # If the first header is not the expected 'Catoragory', assume an index column exists
        if len(header_vals) > 0 and header_vals[0] != 'Catoragory':
            start_col_offset = 1

    # Column letters for MoQ and Qty depend on offset
    moq_col_letter = chr(ord('A') + start_col_offset + 3)  # D normally
    qty_col_letter = chr(ord('A') + start_col_offset + 5)  # F normally

    # Highlight Qty cells where Qty < MoQ
    for row in range(2, total_rows + 1):  # Excel rows start at 1, header is row 1
        moq_val = sheet.range(f"{moq_col_letter}{row}").value
        qty_val = sheet.range(f"{qty_col_letter}{row}").value
        # debug prints (useful to see what's being read)
        print(f"Row {row}: {moq_col_letter}MoQ={moq_val!r}, {qty_col_letter}Qty={qty_val!r}")
        try:
            moq = float(moq_val)
            qty = float(qty_val)
        except (TypeError, ValueError):
            # skip coloring if values can't be interpreted as numbers
            continue
        if qty < moq:
            sheet.range(f"{qty_col_letter}{row}").color = (255, 200, 200)  # Light red

    # Save changes to the open workbook (don't pass filename to avoid SaveAs issues)
    wb.save()
    try:
        wb.set_mock_caller()
    except Exception:
        # set_mock_caller is for testing; ignore if not applicable
        pass
    wb.app.visible = True

def main():
    write_or_open_excel()

if __name__ == "__main__":
    main()
