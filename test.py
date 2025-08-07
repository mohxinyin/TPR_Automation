import win32com.client

try:
    excel = win32com.client.Dispatch("Excel.Application")
    print("✅ Excel COM object created successfully!")
    print(f"Excel Version: {excel.Version}")
    excel.Quit()
except Exception as e:
    print(f"❌ Failed to create Excel COM object: {e}")
