import constants as c
import shutil

from win32com.client import gencache

from file_handler import load_and_convert_csv
from file_handler import load_excel_workbook
from file_handler import open_excel_with_win32
from file_handler import close_excel_with_win32

from worksheet_manager import prepare_working_sheet
from worksheet_manager import adjust_column_width
from worksheet_manager import copy_header_styles
from worksheet_manager import create_summary_sheet
from worksheet_manager import remove_unwanted_columns
from worksheet_manager import create_new_columns
from worksheet_manager import import_inventory_sheet
from worksheet_manager import format_due_date
from worksheet_manager import empty_folder

from helper import fill_schedule_values
from helper import pivot_table_generator
from helper import create_filtered_sheets
from helper import tpr_sheet_config
from helper import convert_to_numeric
from helper import summary_sheet_config

from data_manipulation import fill_blank_due_dates
from data_manipulation import insert_inventory_formula
from data_manipulation import insert_pt
from data_manipulation import create_TPR_columns
from data_manipulation import generate_formula_TPR_SUMMARY

from overview import create_overview_sheet
from overview import copy_paste_as_values

from COMfix import clear_cache

def main():

    clear_cache()
    
######################### USING OPENPYXL ############################

    # Duplicate the excel file and place in another folder  
    shutil.copyfile(c.source_file, c.dest_file)
    shutil.copyfile(c.source_file, c.dest_summary_file)

    # Load workbooks 
    main_wb = load_excel_workbook(c.dest_file)
    summary_wb = load_excel_workbook(c.dest_summary_file)
    header_wb = load_excel_workbook(c.header_file)

    # Rename the first sheet of main file 
    main_wb.active.title = "Sheet 1"

    # Rename the first sheet of summary file 
    summary_wb.active.title = "Sheet 1"

###################### MAIN FILE ##############################################
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

    # Create overview sheet 
    #create_overview_sheet(main_wb)

    # Miscellaneous
    copy_header_styles(working_sheet,main_wb,header_row=1)
    remove_unwanted_columns(MRP_sheet,c.COLUMNS_TO_DELETE_MRP)
    adjust_column_width(main_wb) # Adjust column width so that everything can be seen clearly 
    format_due_date(main_wb,c.due_date_cols) # Format due dates to look like dd/mm/yyyy
    main_wb.save(c.dest_file)

########################### SUMMARY FILE #######################################

    # Prepare TPR Working sheet 
    prepare_working_sheet(summary_wb,header_wb,'TPR Working','SummaryHeader',c.COLUMNS_TO_DELETE_SUMMARY_WORKING) # Prepare Working tab with header 
    tpr_working_sheet = summary_wb['TPR Working']

    # Prepare all filtered sheets ('OHS','MO','SO','PO','Forecast','Suggestion')
    create_filtered_sheets(summary_wb,summary_sheet_config,'TPR Working')
    create_summary_sheet(summary_wb)
    convert_to_numeric(summary_wb)

    # Summary sheet 
    summary_sheet = summary_wb['Summary']
    
    # Miscellaneous
    copy_header_styles(tpr_working_sheet,summary_wb,header_row=1)
    adjust_column_width(summary_wb)
    create_new_columns(summary_sheet,c.COLUMNS_TO_ADD_SUMMARY)
    format_due_date(summary_wb,c.due_date_idx_summary_cols)
    
    summary_wb.save(c.dest_summary_file)

# ######################### USING WIN32 LIB ###############################

    # Open excel TPR and Header wb using win32 
    excel = gencache.EnsureDispatch("Excel.Application")
    try:
        _, wb_main = open_excel_with_win32(excel,c.file_path_win32) # open main file
        _,wb_summary = open_excel_with_win32(excel,c.file_path_summary_win32) # open dest summary file 
        _, wb_header = open_excel_with_win32(excel,c.header_path_win32) # open header file 
    except Exception as e:
        print(f"Failed to open excel workbook: {e}")
        return
    
    # try:
    #     _, wb_header = open_excel_with_win32(c.header_path_win32)
    # except Exception as e:
    #     print(f"Failed to open header workbook: {e}")
    #     return

    # if wb_header is None:
    #     print("Header workbook is None — cannot proceed.")
    #     return

    # Insert 'Inventory by WH' formula
    insert_inventory_formula(wb_main,wb_header)

    # Create pivot tables in 'MRP','Schedule' and 'Inventory by WH' tabs 
    for config in pivot_table_generator():
        insert_pt(wb_main,**config)

    create_TPR_columns(wb_main)

    generate_formula_TPR_SUMMARY(wb_main,'TPR Inventory',c.formula_map_tpr) # Generate formulas for the tpr inventory sheet
    generate_formula_TPR_SUMMARY(wb_main,'Overview',c.formula_map_overview) # Generate formulas for the tpr overview sheet
    generate_formula_TPR_SUMMARY(wb_summary,'Summary',c.formula_map_summary) # Generate formulas for the tpr summary sheet

    copy_paste_as_values(wb_main) # Copy paste the entire overview sheet as values so the formulas dont show 

    # Save and close excel wb 
    close_excel_with_win32(wb_main) 
    close_excel_with_win32(wb_summary)

    # Quit excel processes
    excel.Quit()

    # Empty folder 
    empty_folder()

if __name__ == "__main__":
    main()