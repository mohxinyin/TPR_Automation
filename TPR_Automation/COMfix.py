import shutil
import os

def clear_cache():
    gen_py_path = os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py')

    if os.path.exists(gen_py_path):
        shutil.rmtree(gen_py_path)
        print("✅ gen_py cache cleared.")

