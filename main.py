import constants as c

from file_handler import load_and_convert_csv
from file_handler import load_excel_workbook
from file_handler import open_excel_with_win32
from file_handler import close_excel_with_win32

from worksheet_manager import prepare_working_sheet
from worksheet_manager import adjust_column_width
from worksheet_manager import copy_header_styles
from worksheet_manager import remove_unwanted_columns
from worksheet_manager import create_new_columns
from worksheet_manager import import_inventory_sheet
from worksheet_manager import format_due_date

from helper import fill_schedule_values
from helper import pivot_table_generator
from helper import create_filtered_sheets
from helper import tpr_sheet_config
from helper import convert_to_numeric

from data_manipulation import fill_blank_due_dates
from data_manipulation import insert_inventory_formula
from data_manipulation import insert_pt
from data_manipulation import create_TPR_columns
from data_manipulation import generate_formula_TPR_SUMMARY

def main():

######################### USING OPENPYXL ############################

    # Convert csv to excel file 
    load_and_convert_csv(c.source_file,c.dest_file)

    # Load workbooks 
    main_wb = load_excel_workbook(c.dest_file)
    header_wb = load_excel_workbook(c.header_file)

    # Working tab 
    prepare_working_sheet(main_wb,header_wb,'Working','Header', c.COLUMNS_TO_DELETE_WORKING) # Prepare Working tab with header 
    working_sheet = main_wb['Working']

    # Prepare all filtered sheets 
    create_filtered_sheets(main_wb,tpr_sheet_config,'Working')
    convert_to_numeric(main_wb)

    # MRP tab 
    MRP_sheet = main_wb['MRP']

    # Schedule tab
    schedule_sheet = main_wb['Schedule']
    fill_blank_due_dates(schedule_sheet)
    create_new_columns(schedule_sheet,c.COLUMNS_TO_ADD_SCHEDULE)
    fill_schedule_values(schedule_sheet)

    # Inventory by WH tab 
    import_inventory_sheet(c.qoh_file,main_wb)
    inventory_sheet = main_wb['Inventory by WH']
    create_new_columns(inventory_sheet,c.COLUMN_TO_ADD_WH,'E')

    # Miscellaneous
    copy_header_styles(working_sheet,main_wb,header_row=1)
    remove_unwanted_columns(MRP_sheet,c.COLUMNS_TO_DELETE_MRP)
    adjust_column_width(main_wb) # Adjust column width so that everything can be seen clearly 
    format_due_date(main_wb,c.due_date_idx) # Format due dates to look like dd/mm/yyyy
    main_wb.save(c.dest_file)

######################### USING WIN32 LIB ###############################

    # Open excel TPR and Header wb using win32 
    try:
        excel, wb_main = open_excel_with_win32(c.file_path_win32)
    except Exception as e:
        print(f"Failed to open main workbook: {e}")
        return
    
    try:
        _, wb_header = open_excel_with_win32(c.header_path_win32)
    except Exception as e:
        print(f"Failed to open header workbook: {e}")
        return

    if wb_header is None:
        print("Header workbook is None — cannot proceed.")
        return

    # Insert 'Inventory by WH' formula
    insert_inventory_formula(wb_main,wb_header)

    # Create pivot tables in 'MRP','Schedule' and 'Inventory by WH' tabs 
    for config in pivot_table_generator():
        insert_pt(wb_main,**config)

    create_TPR_columns(wb_main)
    generate_formula_TPR_SUMMARY(wb_main,'TPR Inventory',c.formula_map_tpr) # generate formulas for the tpr inventory and summary sheets

    # Save and close excel wb 
    close_excel_with_win32(excel,wb_main) 

if __name__ == "__main__":
    main()