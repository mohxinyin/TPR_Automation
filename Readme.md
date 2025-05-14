# **TPR and TPR Summary Automator** 

## **Overview** 
This program generates the Time Phase Material Requirement and Time Phase Material Requirement Summary Reports

## **How It Works** 
1. There are 2 main files to be run, main.py (generates TPR Report) and main_summary.py (generates TPR Summary Report)
2. The program will begin by converting the csv file to excel 
3. It will then load all the necessary excel files 
4. It will first generate the **working** sheet 
5. After generating the working sheet, it will filter and sort data into the different sheets (TPR:'MRP','Schedule','Inventory by WH','TPR Inventory', Summary: 'OHS','MO','PO','SO','Summary')
6. All the necessary pivot tables and formulas will be inserted 
7. Once done, both reports will be generated 
8. files required: **Quantity on Hand** file, **TPR** file, **TPR Header** file
9. NOTE: have to use **absoulte** file path for destination files for win32 library  

## **Things to Improve** 
- Sometimes cache will be corrupted, run **COMfix.py** to clear cache Error msg: \[ERROR] Could not open Excel or workbook: module 'win32com.gen\_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'
- Sometimes, COM automation error related to pywin32 will pop up, run **makepy.py** to fix error: [ERROR] Could not open Excel or workbook: This COM object can not automate the makepy process - please run makepy manually for this object
- In the schedule pivot table, the dates are not highlighted, have to manually highlight 
- Excel application will pop up when main_summary.py is run 

## **How to run** 
To run the project, simply execute the following commands:

```bash
python main.py
python main_summary.py
