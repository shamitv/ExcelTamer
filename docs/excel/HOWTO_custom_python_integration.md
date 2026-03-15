# Python Excel Integration HOW-TO Guide

This guide is intended for developers who want to integrate Microsoft Excel automation directly into their own Python codebase. It uses the [xlwings](https://www.xlwings.org/) library, which enables Python to interact with a live Excel application via COM on Windows.

> **Why xlwings?** 
> Unlike `openpyxl` or `pandas` (which read the static file structure), `xlwings` connects to the running Excel app. This allows you to interact with unsaved changes, read evaluated formulas, apply formatting, and capture exact visual screenshots of how data is rendered.

---

## 1. Installation

Install `xlwings` via pip. If you want to use data structures like DataFrames, you will also need `pandas`.

```bash
pip install xlwings pandas
```

---

## 2. Managing the Excel Application & Workbooks

To safely operate Excel, you need to manage the `App` (the process) and the `Book` (the file). 

```python
import xlwings as xw

# Connect to the active Excel application, or start a new visible instance
app = xw.apps.active if xw.apps else xw.App(visible=True)

# 1. Open an existing file
wb = app.books.open("C:/path/to/financials.xlsx")

# 2. Add a new blank workbook
# wb = app.books.add()

# 3. Connect to the active workbook (if one is already open in Excel)
# wb = app.books.active
```

**Clean Up**: Always close workbooks when finished, otherwise Excel processes can hang in the background.

```python
wb.save("C:/path/to/financials.xlsx")  # Save changes
wb.close()                             # Close the workbook
# app.quit()                           # (Optional) Quit Excel entirely
```

---

## 3. Working with Sheets and Cells

### Navigating Sheets

```python
# List all sheet names
sheet_names = [sheet.name for sheet in wb.sheets]
print(sheet_names)

# Select a specific sheet
sheet = wb.sheets["Income Statement"]

# Add or delete sheets
# wb.sheets.add("New Sheet")
# wb.sheets["Old Sheet"].delete()
```

### Reading and Writing Single Cells

`xlwings` automatically converts between Python types (int, float, string, datetime) and Excel types.

```python
# Write a value or formula
sheet.range("B2").value = 50000
sheet.range("B3").value = "=B2 * 1.1"

# Read a value (evaluates formulas automatically)
current_revenue = sheet.range("B3").value
print(current_revenue) # Output: 55000.0

# Read the raw formula string instead of the evaluated value
formula = sheet.range("B3").formula
print(formula) # Output: =B2*1.1
```

---

## 4. Bulk Data operations: `pandas` Integration

Interacting with Excel cell-by-cell via COM is extremely slow. For ranges of data, use `pandas` DataFrames. `xlwings` supports writing and reading DataFrames natively.

### Writing a DataFrame to Excel

```python
import pandas as pd

# Create a sample DataFrame
data = {
    "Category": ["Revenue", "COGS", "Gross Margin"],
    "Q1": [1000, 400, 600],
    "Q2": [1200, 450, 750]
}
df = pd.DataFrame(data)

# Write it starting at cell A1 (index=False hides the pandas row numbers)
sheet.range("A1").options(index=False).value = df
```

### Reading an Excel Range into a DataFrame

```python
# Read a specific range into a DataFrame, using the first row as headers
df_read = sheet.range("A1:C4").options(pd.DataFrame, index=False, header=1).value
print(df_read)

# Read the entire "used range" (all active cells on the sheet)
used_range_df = sheet.used_range.options(pd.DataFrame, index=False, header=1).value
```

---

## 5. Capturing Screenshots of Excel Data

Because `xlwings` uses COM, you can command Excel to render a specific range and save it as an image. This is incredibly useful for capturing charts, formatting, and conditional styling.

```python
import os

sheet = wb.sheets["Dashboard"]

# Select the range you want to capture
target_range = sheet.range("A1:H20")

# 1. Bring Excel to the foreground and scroll the range into view
target_range.api.Show()

# 2. Capture and save as PNG
output_path = os.path.abspath("./dashboard_capture.png")
target_range.to_png(output_path)

print(f"Screenshot saved to {output_path}")
```

> **Tip for taking screenshots of an entire sheet**:
> If you don't know the exact range, you can use `sheet.used_range` to capture all populated cells:
> ```python
> sheet.used_range.api.Show()
> sheet.used_range.to_png(output_path)
> ```

---

## 6. How to Iterate Through All Sheets & Take Screenshots

Combining the concepts above, here is a complete, copy-pasteable script to open a workbook, iterate through every sheet, and save a screenshot of each sheet's used data.

```python
import os
import xlwings as xw

def capture_all_sheets(excel_path: str, output_folder: str):
    # Ensure absolute paths
    excel_path = os.path.abspath(excel_path)
    output_folder = os.path.abspath(output_folder)
    os.makedirs(output_folder, exist_ok=True)
    
    # 1. Connect to Excel and open the workbook
    app = xw.apps.active if xw.apps else xw.App(visible=True)
    wb = app.books.open(excel_path)
    
    try:
        # 2. Iterate through all sheets
        for sheet in wb.sheets:
            print(f"Processing sheet: {sheet.name}...")
            
            # The used_range contains all cells with data or formatting
            data_range = sheet.used_range
            
            # Skip completely empty sheets
            if not data_range.value:
                print(f"  Skipping {sheet.name} (Empty)")
                continue
                
            # 3. Bring into view and capture screenshot
            data_range.api.Show()
            
            # Save the PNG
            output_png = os.path.join(output_folder, f"{sheet.name}.png")
            data_range.to_png(output_png)
            print(f"  Saved -> {output_png}")
            
    finally:
        # 4. Clean up
        wb.close()
        # Optional: app.quit() if you want to force-close the Excel process completely

if __name__ == "__main__":
    capture_all_sheets(
        excel_path="financial_report.xlsx", 
        output_folder="./screenshots"
    )
```

---

## 7. Critical: The COM Threading Model

If you are building an API, an async server (like FastAPI), or a multi-threaded application, you **must** be aware of Excel's threading model.

Microsoft Excel uses the **Single-Threaded Apartment (STA)** COM model. 

**The Golden Rule:** The thread that creates the Excel COM object (`xw.App` / `xw.Book`) is the **only** thread allowed to interact with it.

If your Python app uses tools like `asyncio.run_in_executor()` or `concurrent.futures`, and separate threads try to read/write to the `wb` object, Excel will crash or throw obscure COM `RPC_E_WRONG_THREAD` exceptions.

### The Solution: A Dedicated Worker Thread
Funnel all Excel operations through a single dedicated thread. 

```python
from concurrent.futures import ThreadPoolExecutor

# Create a thread pool with exactly ONE worker
excel_executor = ThreadPoolExecutor(max_workers=1)

# All Excel object creation AND manipulation must be submitted to this executor
future_app = excel_executor.submit(xw.App, visible=True)
app = future_app.result()

def write_data():
    wb = app.books.active
    wb.sheets[0].range("A1").value = "Safe Write"

# Submit work to the dedicated COM thread
excel_executor.submit(write_data).result()
```
By enforcing a strict 1-worker thread queue, you guarantee thread safety with Excel COM objects even in highly concurrent Python web servers.
