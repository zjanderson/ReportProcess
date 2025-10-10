import os
import sys
import win32com.client as win32
# File paths
report_path = r"C:\Users\zanderson\Downloads\Report.xlsx"
pick_output_path = r"C:\Users\zanderson\Downloads\Pick.xlsx"

# Validate source file exists
if not os.path.exists(report_path):
    print(f"❌ Error: Source file not found: {report_path}")
    sys.exit(1)

# Check if output directory exists
output_dir = os.path.dirname(pick_output_path)
if not os.path.exists(output_dir):
    print(f"❌ Error: Output directory does not exist: {output_dir}")
    sys.exit(1)

# Column mapping
column_mapping = {
    'Primary Reference': 'BOL #',
    'Purchase Order Number': 'PO #',
    'Carrier Name': 'Carrier Name',
    'Owner': 'Owner',
    'Origin Name': 'Origin Name',
    'Dest Name': 'Dest Name',
    'Target Ship (Early)': 'Target Ship (Early)',
    'Act Date: Pick Appt': 'Act Date: Pick Appt',
    'Target Delivery (Early)': 'Target Delivery (Early)',
    'Act Date: Drop Appt': 'Act Date: Drop Appt',
    'Missed Pick': 'Missed Pick'
}

# Start Excel with error handling
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
except (OSError, RuntimeError) as e:
    print(f"❌ Error: Failed to start Excel: {e}")
    sys.exit(1)

# Open source workbook with error handling
try:
    wb = excel.Workbooks.Open(report_path)
except (OSError, FileNotFoundError) as e:
    print(f"❌ Error: Failed to open workbook: {e}")
    excel.Quit()
    sys.exit(1)

# Validate and get the worksheet
sheet_name = 'Zac'
try:
    ws = wb.Worksheets(sheet_name)
except (OSError, AttributeError, KeyError) as e:
    print(f"❌ Error: Worksheet '{sheet_name}' not found in workbook")
    print(f"Available worksheets: {[ws.Name for ws in wb.Worksheets]}")
    wb.Close(SaveChanges=False)
    excel.Quit()
    sys.exit(1)

# Get used range
n_rows = ws.UsedRange.Rows.Count
n_cols = ws.UsedRange.Columns.Count

# Get headers with validation
headers = [ws.Cells(1, col).Value for col in range(1, n_cols + 1)]

# Validate required columns exist
required_columns = list(column_mapping.keys())
missing_columns = [col for col in required_columns if col not in headers]

if missing_columns:
    print(f"❌ Error: Required columns not found: {missing_columns}")
    print(f"Available columns: {headers}")
    wb.Close(SaveChanges=False)
    excel.Quit()
    sys.exit(1)

missed_pick_col_idx = headers.index('Missed Pick') + 1

# Determine columns to keep and rename
keep_cols = [col for col in column_mapping.keys()]
keep_col_indices = [headers.index(col) + 1 for col in keep_cols]
renamed_headers = [column_mapping[col] for col in keep_cols]

# Create new workbook with error handling
try:
    pick_wb = excel.Workbooks.Add()
    pick_ws = pick_wb.Worksheets(1)
    pick_ws.Name = 'Pick'
except (OSError, RuntimeError) as e:
    print(f"❌ Error: Failed to create new workbook: {e}")
    wb.Close(SaveChanges=False)
    excel.Quit()
    sys.exit(1)

# Write renamed headers
for col_idx, new_name in enumerate(renamed_headers, start=1):
    pick_ws.Cells(1, col_idx).Value = new_name
    pick_ws.Cells(1, col_idx).Font.Bold = True

# Copy rows without 'Missed Pick' values
target_row = 2
for row in range(2, n_rows + 1):
    missed_pick_value = ws.Cells(row, missed_pick_col_idx).Value
    if missed_pick_value is not None and str(missed_pick_value).strip() != "":
        continue

    for i, src_col_idx in enumerate(keep_col_indices, start=1):
        source_cell = ws.Cells(row, src_col_idx)
        target_cell = pick_ws.Cells(target_row, i)
        target_cell.Value = source_cell.Value
        target_cell.Interior.Color = source_cell.Interior.Color
        target_cell.Font.Color = source_cell.Font.Color
        target_cell.Font.Bold = source_cell.Font.Bold
        target_cell.Font.Italic = source_cell.Font.Italic
        target_cell.Font.Size = source_cell.Font.Size
        target_cell.Borders.LineStyle = source_cell.Borders.LineStyle

    target_row += 1

# Save and close with error handling
try:
    pick_wb.SaveAs(pick_output_path)
    print("✅ Pick.xlsx has been created with renamed columns and formatting preserved.")
except (OSError, PermissionError) as e:
    print(f"❌ Error: Failed to save output file: {e}")
finally:
    # Ensure cleanup happens even if save fails
    try:
        pick_wb.Close(SaveChanges=False)
    except (OSError, AttributeError):
        pass
    try:
        wb.Close(SaveChanges=False)
    except (OSError, AttributeError):
        pass
    try:
        excel.Quit()
    except (OSError, AttributeError):
        pass