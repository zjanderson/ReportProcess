import win32com.client as win32

# File paths
report_path = r"C:\Users\zanderson\Downloads\Report.xlsx"
pick_output_path = r"C:\Users\zanderson\Downloads\Pick.xlsx"

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

# Start Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

# Open source workbook and sheet
wb = excel.Workbooks.Open(report_path)
ws = wb.Worksheets('Zac')

# Get used range
n_rows = ws.UsedRange.Rows.Count
n_cols = ws.UsedRange.Columns.Count

# Get headers
headers = [ws.Cells(1, col).Value for col in range(1, n_cols + 1)]
missed_pick_col_idx = headers.index('Missed Pick') + 1

# Determine columns to keep and rename
keep_cols = [col for col in column_mapping.keys()]
keep_col_indices = [headers.index(col) + 1 for col in keep_cols]
renamed_headers = [column_mapping[col] for col in keep_cols]

# Create new workbook
pick_wb = excel.Workbooks.Add()
pick_ws = pick_wb.Worksheets(1)
pick_ws.Name = 'Pick'

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

# Save and close
pick_wb.SaveAs(pick_output_path)
pick_wb.Close(SaveChanges=False)
wb.Close(SaveChanges=False)
excel.Quit()

print("✅ Pick.xlsx has been created with renamed columns and formatting preserved.")