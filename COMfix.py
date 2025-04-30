import shutil
import os
import win32com.client

gen_py_path = os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py')

if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)
    print("✅ gen_py cache cleared.")

# This line forces regeneration
excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
print("✅ Excel dispatch successful")
