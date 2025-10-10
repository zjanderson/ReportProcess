import pandas as pd
import win32com.client as win32

# File path
report_path = r"C:\Users\zanderson\Downloads\Report.xlsx"

# Load all sheets
xls = pd.read_excel(report_path, sheet_name=None)

# Step 1: Filter rows where 'Name' column contains 'Zac'
filtered_sheets = {}
for sheet_name, df in xls.items():
    if 'Name' in df.columns:
        df_filtered = df[df['Name'].astype(str).str.contains('Zac', na=False)].copy()
    else:
        df_filtered = df.copy()
    filtered_sheets[sheet_name] = df_filtered

# Step 2: Transfer sheet will be empty except for header
if 'Transfer' in filtered_sheets:
    filtered_sheets['Transfer'] = filtered_sheets['Transfer'].iloc[0:0]

# Step 3: Rename columns in Watchlist sheet
watchlist_mapping = {
    'Primary Reference': 'BOL #',
    'Purchase Order Number': 'PO #',
    'Carrier Name': 'Carrier Name',
    'Owner': 'Owner',
    'Origin Name': 'Origin Name',
    'Dest Name': 'Dest Name',
    'Target Ship (Early)': 'Target Ship (Early)',
    'Target Delivery (Early)': 'Target Delivery (Early)',
}
if 'Watchlist' in filtered_sheets:
    df_watchlist = filtered_sheets['Watchlist']
    columns_to_keep = [col for col in watchlist_mapping if col in df_watchlist.columns]
    df_watchlist = df_watchlist[columns_to_keep].rename(columns=watchlist_mapping)
    filtered_sheets['Watchlist'] = df_watchlist

# Step 4: Combine Missed PU and Missed DROP into PickDrop
missed_pu_sheet = None
missed_drop_sheet = None
for name in filtered_sheets:
    if name.lower() in ['missed pu', 'missed pick']:
        missed_pu_sheet = filtered_sheets[name]
    elif name.lower() in ['missed drop', 'missed del']:
        missed_drop_sheet = filtered_sheets[name]

pickdrop_df = pd.DataFrame()
if missed_pu_sheet is not None:
    missed_pu_sheet = missed_pu_sheet.copy()
    if 'Target Ship (Early)' in missed_pu_sheet.columns:
        missed_pu_sheet['Highlight'] = 'Target Ship (Early)'
    pickdrop_df = pd.concat([pickdrop_df, missed_pu_sheet], ignore_index=True)

if missed_drop_sheet is not None:
    missed_drop_sheet = missed_drop_sheet.copy()
    if 'Target Delivery (Early)' in missed_drop_sheet.columns:
        missed_drop_sheet['Highlight'] = 'Target Delivery (Early)'
    pickdrop_df = pd.concat([pickdrop_df, missed_drop_sheet], ignore_index=True)

filtered_sheets['PickDrop'] = pickdrop_df

# Use win32 to overwrite Report.xlsx
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False  # Suppress confirmation dialogs

wb = excel.Workbooks.Open(report_path)

# Add a temporary sheet to avoid Excel error
temp_sheet = wb.Sheets.Add()
temp_sheet.Name = "TempSheet"

# Delete all existing sheets except the temporary one
for sheet in list(wb.Sheets):
    if sheet.Name != "TempSheet":
        sheet.Delete()

# Add new sheets and write data
for sheet_name, df in filtered_sheets.items():
    ws = wb.Sheets.Add()
    ws.Name = sheet_name
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.Cells(1, col_idx).Value = col_name
    for row_idx, row in enumerate(df.values, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.Cells(row_idx, col_idx).Value = value

# Delete the temporary sheet
wb.Sheets("TempSheet").Delete()

# Save and close
wb.Save()
wb.Close(False)
excel.Quit()

print("✅ Report.xlsx has been updated successfully.")