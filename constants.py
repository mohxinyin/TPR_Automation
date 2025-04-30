# CONSTANTS

# Files
source_file = 'source/TPR(sample).csv'
header_file = 'source/TPR HEADER.xlsx'
qoh_file = 'source/QOH.xlsx' # Quantity on hand file 
dest_file = 'dest/TPR(final).xlsx'
dest_summary_file = 'dest/TPR_SUMMARY(final).xlsx'

# Files for win32 library
file_path_win32 = r"C:\Users\xinyi.moh\ExcelAutomation\dest\TPR(final).xlsx"
file_path_summary_win32 = r"C:\Users\xinyi.moh\ExcelAutomation\dest\TPR_SUMMARY(final).xlsx"
header_path_win32 = r"C:\Users\xinyi.moh\ExcelAutomation\source\TPR HEADER.xlsx"

# Columns to be deleted 
COLUMNS_TO_DELETE_WORKING  = [
    "A",  # Single column
    *[chr(c) for c in range(ord("C"), ord("K"))],  # C to J
    "O", "P",
    *[chr(c) for c in range(ord("S"), ord("U"))] + ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT'],
    "AV",
    *["BD", "BE", "BF", "BG", "BH", "BI", "BJ"]
]
COLUMNS_TO_DELETE_SUMMARY_WORKING = [
    "A",  # Single column
    *[chr(c) for c in range(ord("C"), ord("K"))],  # C to J
    "O", "P", "R",  # Individual columns
    *[chr(c) for c in range(ord("T"), ord("Z") + 1)],  # T to Z
    *["AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT"],  # AC to AT
    *[f"B{chr(c)}" for c in range(ord("D"), ord("K"))]  # BD to BJ
]
COLUMNS_TO_DELETE_MRP = ['L', 'M', 'N', 'O','P']
COLUMNS_TO_DELETE_SCHEDULE = ['U','V']

# Columns to be added 
COLUMNS_TO_ADD_SCHEDULE = ['MRP','MO','EXPEDITE','POSTPONE']
COLUMN_TO_ADD_WH = ['Area']
COLUMNS_TO_ADD_TPR = ['Delta','Total','Delta2']
COLUMNS_TO_COPY_SUMMARY = [1, 2, 3, 4, 5, 6, 7, 16] # Columns A to G and P 
COLUMNS_TO_ADD_SUMMARY = ['MO','PO','SO','Forecast','MO Comp','Available','Available with MRP','MRP','MRP Comp','Suggestion','Demand','Supply']

# Table ranges 
MRP_Table_Range = 'MRP!$A:$K' 
Schedule_Table_Range = 'Schedule!$A:$V'
Inventory_Pivot_Range = "'Inventory by WH'!$B:$H"
Inventory_Formula_Range = 'F:F'

# Due Date , year and month column indexes for schedule sheet  
due_date_idx = 10 

#TPR formula 
formula_map_tpr = {
    '17': "=IFERROR(VLOOKUP($A{row}, 'Inventory by WH'!$O:$T,2,FALSE),0)", # Column Q
    '18': "=IFERROR(VLOOKUP($A{row}, 'Inventory by WH'!$O:$T,3,FALSE),0)", # Column R
    '19': "=IFERROR(VLOOKUP($A{row}, 'Inventory by WH'!$O:$T,4,FALSE),0)", # Column S
    '20': "=IFERROR(VLOOKUP($A{row}, 'Inventory by WH'!$O:$T,5,FALSE),0)", # Column T
    '21': "=IFERROR(VLOOKUP($A{row}, 'Inventory by WH'!$O:$T,6,FALSE),0)", # Column U
    '22': "=IFERROR(--(U{row}=K{row}), FALSE)", # Column V
    '23': "=SUM(Q{row}:U{row})", # Column W
    '24': "=IFERROR(--(K{row}=W{row}), FALSE)" # Column X
}

#TPR Summary formula 
formula_map_summary = {
    '9': '=IFERROR(SUMIFS(MO!$N:$N,MO!$A:$A,Summary!$A{row},MO!$L:$L,"<>Job: MRP*"),0)', # Column I (MO)
    '10': "=IFERROR(SUMIFS(PO!N:N,PO!A:A,Summary!A{row}),0)", # Column J (PO)
    '11': "=IFERROR(SUMIFS(SO!O:O,SO!A:A,Summary!A{row}),0)", # Column K (SO)
    '12': "=IFERROR(SUMIFS(Forecast!O:O,Forecast!A:A,Summary!A{row}),0)", # Column L (Forecast)
    '13': '=IFERROR(SUMIFS(MO!$O:$O,MO!$A:$A,Summary!$A{row},MO!$L:$L,"<>Job: MRP*"),0)', # Column M (MO Comp)
    '14': "=H{row}+I{row}+J{row}-K{row}-L{row}-M{row}", # Column N (Available)
    '15': "=N{row}+P{row}+R{row}-Q{row}", # Column (Available with MRP)
    '16': '=IFERROR(SUMIFS(MO!$N:$N,MO!$A:$A,Summary!$A{row},MO!$L:$L,"=Job: MRP*"),0)' ,# Column P (MRP)
    '17': '=IFERROR(SUMIFS(MO!$O:$O,MO!$A:$A,Summary!$A{row},MO!$L:$L,"=Job: MRP*"),0)', # Column Q (MRP Comp)
    '18': "=IFERROR(SUMIFS(Suggestion!N:N,Suggestion!A:A,Summary!A{row}),0)", # Column R (Suggestion)
    '20': "=(K{row}+M{row}+L{row}+Q{row})", # Column T (Demand)
    '21': "=H{row}+J{row}", # Column U (Supply)
    '22': "=U{row}-T{row}", # Column V
    '23': "=IF(O{row}>0,TRUE,0)", # Column W
    '24': '=IF(R{row}>0, TRUE, "")', # Column X 
    '25': "=W{row}=X{row}" # Column X
}

# Define sheets and columns you need to fix
sheets_and_columns = {
    "MO": ["A", "N", "O"],
    "PO": ["A", "N"],
    "SO": ["A", "O"],
    "Forecast": ["A", "O"],
    "Suggestion": ["A", "N"],
}

