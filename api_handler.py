import os
import base64
import requests
from dotenv import load_dotenv

load_dotenv()

def fetch_standard_cost_from_epicor(part_num):
    """
    Fetches standard cost for a PartNum from the Epicor API.
    """
    #Load credentials
    username = os.getenv("EPICOR_API_USERNAME")
    password = os.getenv("EPICOR_API_PASSWORD")
    api_key = os.getenv("EPICOR_API_Key")

    # Encode Basic Auth
    credentials = f"{username}:{password}"
    encoded_credentials = base64.b64encode(credentials.encode()).decode()

    # Construct your API URL
    base_url = f"https://seasiadtapp08.epicorsaas.com/saas583/api/v2/odata/8001/Erp.BO.PartCostSearchSvc/PartCostSearches" 
    # base_url = f"https://seasiadtpilot08.epicorsaas.com/saas583pilot/api/v2/odata/8001/Erp.BO.PartCostSearchSvc/PartCostSearches" #pilot url
    params = {
        "$select": "PartNum,StdTotalCost" ,
        "$filter": f"PartNum eq '{part_num}'",
    }

    headers = {
        "accept": "application/json",
        "Authorization": f"Basic {encoded_credentials}",
        "X-API-Key": api_key
    }

    try:
        response = requests.get(base_url, headers=headers, params=params)
        print("Requesting URL:", base_url)
        print("Headers:", headers)
        response.raise_for_status()
        data = response.json()

        if data.get("value"):
            # Extract 'StandardCost' or the actual field you need
            return data['value'][0].get("StdTotalCost", "Not Found")
        else:
            return "Not Found"

    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for {part_num}: {e}")
        return "Error"

def update_excel_with_standard_cost(wb,input_excel):
    """
    Uses openpyxl to update 'Std Cost' column in 'Overview' sheet using API results.
    """
    ws = wb['Overview']

    # Find column indexes
    partnum_col = None
    std_cost_col = None

    # Assume header is in the first row
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col).value
        if header == 'PartNum':
            partnum_col = col
        if header == 'Std Cost':
            std_cost_col = col

    if partnum_col is None:
        raise ValueError("No 'PartNum' column found.")
    
    if std_cost_col is None:
        raise ValueError("No 'Std Cost' column found.")

    # Loop through rows and update standard cost
    for row in range(2, ws.max_row + 1):
        partnum = ws.cell(row=row, column=partnum_col).value
        if partnum:
            std_cost = fetch_standard_cost_from_epicor(partnum)
            ws.cell(row=row, column=std_cost_col).value = std_cost

    #wb.save(input_excel)
    print(f"\nUpdated 'Std Cost' column in '{input_excel}' successfully.")

