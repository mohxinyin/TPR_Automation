from filtering import filter_and_create_sheet
from filtering import fill_column_based_on_filter
import constants as c 

def create_filtered_sheets(wb,sheet_config, source_sheet_name): # Helper function to filter and create sheets 
    for config in sheet_config():
        filter_and_create_sheet(
            wb=wb,
            filters=config["filters"],
            output_sheet_name=config["name"],
            source_sheet_name = source_sheet_name 
        )

def tpr_sheet_config():
    yield {
        "name": "MRP",
        "filters": [
            lambda df: df['Source'].str.contains('MRP', na=False, case=False),
            lambda df: df['Due Date'].dt.year == 2025,
            lambda df: df['Receipts'].notna() & (df['Receipts'].str.strip() != '')
        ]
    }
    yield {
        "name": "Schedule",
        "filters": [
            lambda df: df['Source'].str.contains('Job', na=False, case=False),
            lambda df: df['Type'] == 'M',
            lambda df: df['Receipts'].notna() & (df['Receipts'].str.strip() != '')
        ]
    }
    yield {
        "name": "TPR Inventory",
        "filters": [
            lambda df: df['Source'].str.contains('On-Hand Quantity', na=False, case=False)
        ]
    }
    print("Filtered sheets created")

def summary_sheet_config():
    yield {
        "name": "OHS",
        "filters": [
            lambda df: df['Source'].str.contains('On-hand', na=False, case=False),
        ]
    }
    yield {
        "name": "MO",
        "filters": [
            lambda df: df['Source'].str.startswith('Job', na=False,)
        ]
    }
    yield {
        "name": "SO",
        "filters": [
            lambda df: df['Source'].str.contains('SO:', na=False, case=False)
        ]
    }
    yield {
        "name": "PO",
        "filters": [
            lambda df: df['Source'].str.contains('PO:', na=False, case=False),
        ]
    }
    yield {
        "name": "Forecast",
        "filters": [
            lambda df: df['Source'].str.contains('Forecast', na=False, case=False),
        ]
    }
    yield {
        "name": "Suggestion",
        "filters": [
            lambda df: df['Source'].str.contains('Suggestion', na=False, case=False),
        ]
    }
    print("Filtered sheets created")

def fill_schedule_values(ws):
    # MRP
    fill_column_based_on_filter(ws,'Q', lambda val: 'MRP' in val.upper()) # Start from Q because of the 2 extra columns added ("year" and "month" columns) in schedule tab
    # MO
    fill_column_based_on_filter(ws,'R', lambda val: val.startswith('Job') and 'MRP' not in val.upper())
    # Expedite
    fill_column_based_on_filter(ws,'S', lambda val: 'expedite' in val.lower())
    # Postpone
    fill_column_based_on_filter(ws,'T', lambda val: 'postpone' in val.lower())

def pivot_table_generator():
    yield{
        'sheet_name' : 'MRP',
        'table_range' : c.MRP_Table_Range,
        'pivot_table_location' :'O1',
        'row_field' : ['Class'],
        'data_field' : [('PartNum','count')]
    }
    yield{
        'sheet_name' : 'Schedule',
        'table_range' : c.Schedule_Table_Range,
        'pivot_table_location' : 'X1',
        'row_field' : ['Year','Month','Due Date'],
        'data_field' : [('MRP','count'),('MO','count'),('EXPEDITE','count'),('POSTPONE','count')],
        'filter_field' : ['Class']
    }
    yield{
        'sheet_name' : 'Inventory by WH',
        'table_range' : c.Inventory_Pivot_Range,
        'pivot_table_location' :'O1',
        'row_field' : ['Part Num'],
        'column_field': ['Area'],
        'data_field' : [('On Hand','sum')]
    }        

def convert_to_numeric(wb):
    """
    Loops through all sheets in the workbook and converts column values 
    to numeric types where possible, excluding specific columns.

    Parameters:
    - wb: openpyxl Workbook object
    """
    skip_columns = ['PartNum','Class','Due Date']

    for sheet_name in wb.sheetnames:
        if sheet_name == 'Sheet1':
            print(f"Skipping sheet: {sheet_name}")
            continue

        ws = wb[sheet_name]
        print(f"Processing sheet: {sheet_name}")

        # Get headers from first row
        headers = [cell.value for cell in ws[1]]

        for col_idx, col_name in enumerate(headers, start=1):
            if col_name not in skip_columns:
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    try:
                        if cell.value not in [None, '']:
                            val = str(cell.value).replace(',', '').strip()
                            cell.value = float(val)
                    except ValueError:
                        print (f"Unable to convert {cell.value} to numeric")

    print(f"All sheets updated with numeric conversions (excluding {skip_columns}).")








