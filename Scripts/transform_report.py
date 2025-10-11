import pandas as pd
import win32com.client as win32

# File path
report_path = r"C:\Users\zanderson\Downloads\Report.xlsx"

# Load all sheets
xls = pd.read_excel(report_path, sheet_name=None)

# Step 1: Filter rows where 'Name' column contains 'Zac'
filtered = {}
for sheet_name, df in xls.items():
    if 'Name' in df.columns:
        filtered[sheet_name] = df[df['Name'].astype(str).str.contains('Zac', na=False)].copy()

# Step 2: Apply yellow fill to Missed PU and Missed DROP before combining
highlight_cells = {}
missed_pu = pd.DataFrame()
missed_drop = pd.DataFrame()

if 'Missed PU' in filtered:
    missed_pu = filtered['Missed PU'].copy()
    if 'Target Ship (Early)' in missed_pu.columns:
        col_idx = missed_pu.columns.get_loc('Target Ship (Early)') + 1
        for i in range(len(missed_pu)):
            highlight_cells[('Missed PU', i + 2, col_idx)] = True

if 'Missed DROP' in filtered:
    missed_drop = filtered['Missed DROP'].copy()
    if 'Target Delivery (Early)' in missed_drop.columns:
        col_idx = missed_drop.columns.get_loc('Target Delivery (Early)') + 1
        for i in range(len(missed_drop)):
            highlight_cells[('Missed DROP', i + 2, col_idx)] = True

# Step 3: Combine into PickDrop
pickdrop = pd.concat([missed_pu, missed_drop], ignore_index=True)

# Step 4: Rename and reorder PickDrop columns
pickdrop_mapping = {
    'LoadID': 'BOL #',
    'Purchase Order Number': 'PO #',
    'PRO Number': 'PRO #',
    'Carrier Name': 'Carrier Name',
    'Owner': 'Owner',
    'Target Ship (Early)': 'Ship Date',
    'Origin Name': 'Origin Name',
    'Origin City': 'Origin City',
    'Origin State': 'State',
    'Target Delivery (Early)': 'Del Date',
    'Dest Name': 'Dest Name',
    'Dest City': 'Dest City',
    'Dest State': 'State',
    'Load note message': 'Load note message'
}
pickdrop_order = list(pickdrop_mapping.values())

pickdrop = pickdrop[[col for col in pickdrop_mapping if col in pickdrop.columns]]
pickdrop = pickdrop.rename(columns=pickdrop_mapping)
pickdrop = pickdrop[pickdrop_order]
filtered['PickDrop'] = pickdrop

# Step 5: Format Watchlist
watchlist_mapping = {
    'Primary Reference': 'BOL #',
    'Purchase Order Number': 'PO #',
    'Carrier Name': 'Carrier Name',
    'Owner': 'Owner',
    'Origin Name': 'Origin Name',
    'Dest Name': 'Dest Name',
    'Target Ship (Early)': 'Target Ship (Early)',
    'Target Delivery (Early)': 'Target Delivery (Early)'
}
watchlist_order = list(watchlist_mapping.values())

if 'Watchlist' in filtered:
    df_watchlist = filtered['Watchlist']
    df_watchlist = df_watchlist[[col for col in watchlist_mapping if col in df_watchlist.columns]]
    df_watchlist = df_watchlist.rename(columns=watchlist_mapping)
    df_watchlist = df_watchlist[watchlist_order]
    filtered['Watchlist'] = df_watchlist

# Step 6: Transfer sheet header only
if 'Transfer' in filtered:
    filtered['Transfer'] = filtered['Transfer'].iloc[0:0]

# Step 7: Write to Excel using win32com
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(report_path)

# Add temp sheet
temp = wb.Sheets.Add()
temp.Name = "TempSheet"

# Delete all other sheets
for sheet in list(wb.Sheets):
    if sheet.Name != "TempSheet":
        sheet.Delete()

# Add new sheets in order
for sheet_name in ['PickDrop', 'Transfer', 'Watchlist']:
    df = filtered.get(sheet_name, pd.DataFrame())
    ws = wb.Sheets.Add()
    ws.Name = sheet_name

    # Write headers
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.Cells(1, col_idx).Value = col_name

    # Write data
    for row_idx, row in enumerate(df.values, start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.Cells(row_idx, col_idx)
            cell.Value = value

# Apply yellow fill to Missed PU and Missed DROP columns before combining
for sheet_name in ['Missed PU', 'Missed DROP']:
    if sheet_name in filtered:
        df = filtered[sheet_name]
        ws = wb.Sheets.Add()
        ws.Name = sheet_name
        for col_idx, col_name in enumerate(df.columns, start=1):
            ws.Cells(1, col_idx).Value = col_name
        for row_idx, row in enumerate(df.values, start=2):
            for col_idx, value in enumerate(row, start=1):
                cell = ws.Cells(row_idx, col_idx)
                cell.Value = value
                if (sheet_name, row_idx, col_idx) in highlight_cells:
                    cell.Interior.Color = 65535  # Yellow

# Delete temp sheet
wb.Sheets("TempSheet").Delete()
wb.Save()
wb.Close(False)
excel.Quit()

print("✅ Report.xlsx updated with PickDrop, Transfer, and Watchlist.")