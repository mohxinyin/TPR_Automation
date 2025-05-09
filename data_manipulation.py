import constants as c 
from worksheet_manager import add_year_month_columns
from datetime import datetime 

def add_pivot_field(pivot_table, field, orientation, position):
    """Helper function to add a field to the pivot table with specified orientation and position."""
    pf = pivot_table.PivotFields(field)
    pf.Orientation = orientation
    pf.Position = position

def pivot_table_config(ws, table_range, pivot_table_location, row_field=None, column_field = None, data_field=None, filter_field=None):
    wb = ws.Parent # Get the workbook from worksheet

    AGGREGATION_FUNCTIONS = {
        'sum': -4157,       # xlSum
        'count': -4112,     # xlCount
    }

    pivot_cache = wb.PivotCaches().Create(
        SourceType=1,  # 1 = xlDatabase
        SourceData=table_range,
        Version=6
    )
    
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=ws.Range(pivot_table_location),
        TableName="PivotTable1"
    )

    # Add Filter Fields (PageField = 3)
    if filter_field:
        for idx, field in enumerate(filter_field, start=1):
            add_pivot_field(pivot_table, field, 3, idx)  # xlPageField

    # Add Row Fields (RowField = 1)
    if row_field:
        for idx, field in enumerate(row_field, start=1):
            add_pivot_field(pivot_table, field, 1, idx)  # xlRowField

    # Column Field (ColumnField = 2)
    if column_field:
        for idx, field in enumerate(column_field, start=1):
            add_pivot_field(pivot_table, field, 2, idx)  # xlColumnField

    # Add data fields with aggregation
    for field_name, agg_type in data_field:
        pf = pivot_table.PivotFields(field_name)

        # Check if 'sum' is requested for a non-numeric column (safe fallback)
        try:
            agg_func_code = AGGREGATION_FUNCTIONS.get(agg_type.lower(), -4112)
            pivot_table.AddDataField(pf, f"{agg_type.capitalize()} of {field_name}", agg_func_code)
        except Exception as e:
            print(f"Failed to add '{agg_type}' for '{field_name}': {e} — Falling back to count.")
            agg_func_code = AGGREGATION_FUNCTIONS['count']
            pivot_table.AddDataField(pf, f"Count of {field_name}", agg_func_code)
    
    if ws.Name == 'Schedule':
        # Get the PivotField for Class
        class_field = pivot_table.PivotFields("Class")

        # Filter the pivot table to show only the 01 and 41 classes 
        allowed_classes = ["01", "41"]

        # Loop through all items in the Class field
        for item in class_field.PivotItems():
            if item.Name in allowed_classes:
                item.Visible = True
            else:
                try:
                    item.Visible = False
                except:
                    print(f"Could not hide item: {item.Name}")    
                    
        # Get current year and month
        current_year = str(datetime.now().year)
        current_month = datetime.now().strftime("%B")  # Example: 'May'

        try:
            pivot_table.ManualUpdate = True

            # Collapse/Expand 'Year' field
            year_field = pivot_table.PivotFields("Year")
            for year_item in year_field.PivotItems():
                year_item.ShowDetail = (year_item.Name == current_year)  # Expand only current year

            # Collapse/Expand 'Month' field
            month_field = pivot_table.PivotFields("Month")
            for month_item in month_field.PivotItems():
                month_item.ShowDetail = (month_item.Name == current_month)  # Expand only current month

            pivot_table.ManualUpdate = False
        except Exception as e:
            print(f"Error collapsing/expanding pivot levels: {e}")

    if ws.Name == 'Inventory by WH':
        try:
            column_labels = pivot_table.PivotFields("Area")

            # Define items you want to hide
            unwanted_columns = ["0", "#N/A", "(blank)"]

            for item in column_labels.PivotItems():
                try:
                    if item.Name in unwanted_columns:
                        item.Visible = False
                    else:
                        item.Visible = True
                except Exception as e:
                    print(f"Could not change visibility for item '{item.Name}': {e}")
        except Exception as e:
            print(f"Error accessing the Column Labels: {e}")

    return pivot_table

def write_summary_info(ws,start_cell = 'O1'): # Write summary info for the MRP sheet (Total No. of Parts, Data shown up till...)
    # Get last row of pivot table (assuming it starts at O1)
    start_row = ws.Range(start_cell).Row
    start_col = ws.Range(start_cell).Column
    cur_row = start_row

    # Find the bottom of pivot table by scanning downward
    while ws.Cells(cur_row, start_col).Value not in [None, ""]:
        cur_row += 1
    summary_row = cur_row + 1  # two rows below the pivot table

    # Count unique parts in Column A
    unique_parts = set()
    last_data_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
    for i in range(2, last_data_row + 1):
        val = ws.Cells(i, 1).Value
        if val:
            unique_parts.add(val)
    ws.Cells(summary_row, start_col).Value = "Total No. of Parts"
    ws.Cells(summary_row, start_col + 1).Value = len(unique_parts)

     # Bold and color the two cells
    for col_offset in range(2):
        cell = ws.Cells(summary_row, start_col + col_offset)
        cell.Font.Bold = True
        cell.Interior.Color = 15773696 # Blue colour     

    largest_month = None
    due_date_col = c.due_date_idx
    last_row = ws.Cells(ws.Rows.Count, due_date_col).End(-4162).Row  # xlUp = -4162 to find the last row with data
    for i in range(2, last_row + 1):
        val = ws.Cells(i, due_date_col).Value  # Column I

        if isinstance(val, datetime):
            if largest_month is None or val > largest_month:
                largest_month = val
        #print(ws.Cells(i, 9).Value, type(ws.Cells(i, 9).Value))

    # Write largest month to sheet
    if largest_month: 
        ws.Cells(summary_row + 2, start_col).Value = f"Data shown up till {largest_month.strftime('%b %Y')}"
        ws.Cells(summary_row + 2, start_col).Font.Bold = True
    
def insert_pt(wb,sheet_name, table_range,pivot_table_location,row_field = None ,column_field = None,data_field = None ,filter_field = None ):

    ws = wb.Sheets(sheet_name)

    if sheet_name.strip() == 'Schedule':
        add_year_month_columns(ws)
        pivot_table_config(ws, table_range, pivot_table_location, row_field, column_field, data_field, filter_field)
        for col_idx in sorted(c.COLUMNS_TO_DELETE_SCHEDULE, reverse=True):
            ws.Columns(col_idx).Delete()
        write_legend(ws)

    if sheet_name.strip() == 'MRP':
        pivot_table_config(ws, table_range, pivot_table_location, row_field, column_field, data_field, filter_field)
        write_summary_info(ws, pivot_table_location) # write summary info if its MRP sheet 
        print("Summary info written in mrp sheet.")

    if sheet_name.strip() == "Inventory by WH":
        ws.Columns(1).Delete()
        col_H = ws.Cells(1,8)
        print(col_H)
        print(type(ws.Range("H2").Value))
        pivot_table_config(ws, table_range, pivot_table_location, row_field, column_field, data_field, filter_field)

    print("Pivot table inserted.")

def fill_blank_due_dates(ws, due_date_col=c.due_date_idx, replacement_date=datetime(2030, 12, 31)):
    """
    Replace blank/empty-looking cells in the due date column with 31/12/2030 in the schedule tab.
    """
    updated = False
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=due_date_col, max_col=due_date_col):
        cell = row[0]
        cell_val = cell.value

        # Check for any kind of blank: None, "", string with only whitespace, Excel-blank, etc.
        if cell_val is None or str(cell_val).strip() == '' or str(cell_val) in ['NaT', 'nan']:
            cell.value = replacement_date
            updated = True

    if updated:
        print("Blanks have been set to 31/12/2030 in the schedule tab.")
    else:
        print("No empty cells found in due date column.")
        print(f"Row {cell.row}: '{cell_val}' (type: {type(cell_val)})")

def write_legend(ws): 
    """
    Writes 'Legend' (bold) and 'Overdue/Late' (italic) beside the pivot table in the schedule sheet.
    """
    start_col = 21  # Column U
    start_row = 1

    # Find the last used row and column from U1 onwards
    first_used_row = start_row
    last_used_col = start_col

    used_range = ws.UsedRange
    max_row = used_range.Rows.Count
    max_col = used_range.Columns.Count

    for row in range(start_row, max_row + 1):
        for col in range(start_col, max_col + 1):
            cell_value = ws.Cells(row, col).Value
            if cell_value not in (None, ""):
                first_used_row = min(first_used_row, row)
                last_used_col = max(last_used_col, col)

    legend_row = first_used_row + 2
    legend_col = last_used_col + 2

    # Write "Legend" (bold)
    legend_cell = ws.Cells(legend_row, legend_col)
    legend_cell.Value = "Legend"
    legend_cell.Font.Bold = True
    legend_cell.Font.Italic = True

    # Write "Overdue/Late" (italic)
    overdue_cell = ws.Cells(legend_row + 1, legend_col)
    overdue_cell.Value = "Overdue/Late"
    overdue_cell.Font.Italic = True
    overdue_cell.Interior.Color = 65535  # Yellow

    print(f"Legend written at {legend_row},{legend_col}")

def insert_inventory_formula(wb_main,wb_header):

    ws = wb_main.Sheets('Inventory by WH')
    header_wb_name = wb_header.Name  # e.g. 'TPR HEADER.xlsx'

    last_row = ws.Cells(ws.Rows.Count, 'E').End(-4162).Row  # -4162 is xlUp
    formula_range = f'F2:F{last_row}'
    formula = f"=VLOOKUP(E2,'[{header_wb_name}]Area'!$A:$B,2,FALSE)" 

    ws.Range(formula_range).Formula = formula
    ws.Range(formula_range).Copy()
    ws.Range(formula_range).PasteSpecial(Paste=-4163)  # Paste as values

    print(f"Area formula(Inventory by WH) pasted")

def create_TPR_columns(wb):

    TPR_sheet = "TPR Inventory"
    Inventory_sheet = "Inventory by WH"

    try:
        # wb = excel.Workbooks.Open(c.file_path_win32)
        source_ws = wb.Sheets(Inventory_sheet)
        target_ws = wb.Sheets(TPR_sheet)

        pivot_table = source_ws.PivotTables(1)
        
        # Get the PivotField used in the Column Labels
        column_field = pivot_table.PivotFields("Area")

        # Extract column label values
        labels = []
        for item in column_field.PivotItems():
            try:
                if item.Visible:
                    labels.append(item.Name)
            except:
                continue
            
        # Find last filled column in Row 1 of target sheet
        last_col = target_ws.Cells(1, target_ws.Columns.Count).End(-4159).Column  # -4159 = xlToLeft

        # Write labels to the next available column block
        for idx, label in enumerate(labels):
            target_ws.Cells(1, last_col + idx + 1).Value = label  # +1 to start at next blank column

        # Append "Delta", "Total", "Delta2" after the labels
        offset = len(labels)
        for i, header in enumerate(c.COLUMNS_TO_ADD_TPR):
            cell = target_ws.Cells(1, last_col + offset + i + 1)
            cell.Value = header
            cell.Font.Bold = True  # Bold the header

        print("Column labels appended successfully.")

    except Exception as e:
        print(f"Error: {e}")

def generate_formula_TPR_SUMMARY(wb, sheet_name, formula_map): # Generate formulas for TPR and Summary reports 

        ws = wb.Sheets(sheet_name)

        # Get the last row using the Excel COM method (win32com)
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp

        # Loop through rows and columns to set the formulas
        for row in range(2, last_row + 1):
            for col_index_str, formula_template in formula_map.items():
                col_index = int(col_index_str)  # Convert string key to integer column index
                formula = formula_template.format(row=row)
                try:
                    ws.Cells(row, col_index).Formula = formula
                except Exception as e:
                    print(f"[ERROR] Failed to insert formula at row {row}, col {col_index}: {formula}")
                    raise

        print("Formulas pasted successfully.")














