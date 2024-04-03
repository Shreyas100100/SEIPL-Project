import subprocess
import sys
import tkinter as tk
from tkinter import filedialog

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def check_and_install_modules():
    required_modules = ["cProfile", "tkinter", "pandas", "customtkinter", "openpyxl"]
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            print(f"{module} is not installed. Installing...")
            install(module)
            print(f"{module} has been installed.")

def choose_excel_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if file_path:
        with open("path.txt", "w") as f:
            f.write(file_path)
            print("Path saved to path.txt:", file_path)
    else:
        print("No file selected.")

if __name__ == "__main__":
    check_and_install_modules()
    choose_excel_path()
