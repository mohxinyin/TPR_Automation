# **TPR and TPR Summary Automator** 

## **Overview** 
This program generates the Time Phase Material Requirement and Time Phase Material Requirement Summary Reports

## **How It Works** 
1. There are **2** main files to be run, main.py (generates TPR Report) and main_summary.py (generates TPR Summary Report)
2. The program will begin by converting the csv file to excel 
3. It will then load all the necessary excel files 
4. It will first generate the **working** sheet 
5. After generating the working sheet, it will filter and sort data into the different sheets (TPR:'MRP','Schedule','Inventory by WH','TPR Inventory', Summary: 'OHS','MO','PO','SO','Summary')
6. All the necessary pivot tables and formulas will be inserted 
7. Once done, both reports will be generated in the destination folder and the source folder will be emptied
8. files required: **Quantity on Hand** file, **TPR** file(source file), **TPR Header** file (header file)
9. NOTE: have to use **absoulte** file path for destination files for win32 library  

## **How to run** 
To run the project, simply:
1. Drop the **Time Phase Requirement** and **Quantity on Hand** files in the **source** folder 
2. Rename the respective files to **'tpr'** and **'qoh'** 
3. Run the main file (**main.py**) followed by summary file (**main_summary.py**)
4. The output files can be found in the **'dest'** folder as **'tpr'** and **'tpr Summary'**

```bash
python main.py 
python main_summary.py 
```

## **Things to Note** 
- Pivot table for 'Inventory by WH' is placed in another cell(shifted from 'O1' to 'S1'), the formula used in 'TPR Inventory' will change according to where columns 'Row Labels' to 'WH' is ($O:$X is changed to $S:$X) due to the increased number of columns in the quantity on hand file 

## **Things to Improve** 
- Sometimes cache will be corrupted, run **COMfix.py** to clear cache Error msg: **[ERROR] Could not open Excel or workbook: module 'win32com.gen\_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'**
- Sometimes, COM automation error related to pywin32 will pop up, run **makepy.py** to fix error: **[ERROR] Could not open Excel or workbook: This COM object can not automate the makepy process - please run makepy manually for this object**
