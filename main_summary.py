import constants as c

from worksheet_manager import prepare_working_sheet
from worksheet_manager import adjust_column_width
from worksheet_manager import copy_header_styles
from worksheet_manager import create_summary_sheet
from worksheet_manager import create_new_columns
from worksheet_manager import format_due_date

from file_handler import load_and_convert_csv
from file_handler import load_excel_workbook
from file_handler import open_excel_with_win32
from file_handler import close_excel_with_win32

from helper import create_filtered_sheets
from helper import summary_sheet_config
from helper import convert_to_numeric

from data_manipulation import generate_formula_TPR_SUMMARY

def main_summary():

######################### USING OPENPYXL ########################

    # Convert csv to excel file 
    load_and_convert_csv(c.source_file,c.dest_summary_file)

    # Load workbooks 
    main_wb = load_excel_workbook(c.dest_summary_file)
    header_wb = load_excel_workbook(c.header_file)

    # Prepare TPR Working sheet 
    prepare_working_sheet(main_wb,header_wb,'TPR Working','SummaryHeader',c.COLUMNS_TO_DELETE_SUMMARY_WORKING) # Prepare Working tab with header 
    tpr_working_sheet = main_wb['TPR Working']

    # Prepare all filtered sheets ('OHS','MO','SO','PO','Forecast','Suggestion')
    create_filtered_sheets(main_wb,summary_sheet_config,'TPR Working')
    create_summary_sheet(main_wb)
    convert_to_numeric(main_wb)

    # Summary sheet 
    summary_sheet = main_wb['Summary']
    
    # Miscellaneous
    copy_header_styles(tpr_working_sheet,main_wb,header_row=1)
    adjust_column_width(main_wb)
    create_new_columns(summary_sheet,c.COLUMNS_TO_ADD_SUMMARY)
    format_due_date(main_wb,c.due_date_idx_summary)


    main_wb.save(c.dest_summary_file)

######################### USING WIN32 LIB ########################

    # Open excel wb using win32 
    excel,wb = open_excel_with_win32(c.file_path_summary_win32)
    summary_sheet = wb.Sheets("Summary")

    generate_formula_TPR_SUMMARY(wb,'Summary',c.formula_map_summary)

    # Save and close excel wb 
    close_excel_with_win32(excel,wb)

if __name__ == "__main__":
    main_summary()








