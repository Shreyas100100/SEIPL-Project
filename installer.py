import subprocess
import sys
import tkinter as tk
from tkinter import filedialog
import requests

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def check_and_install_modules():
    required_modules = ["cProfile", "tkinter", "pandas", "openpyxl"]
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            print(f"{module} is not installed. Installing...")
            install(module)
            print(f"{module} has been installed.")

def download_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"{filename} downloaded successfully.")
    else:
        print(f"Failed to download {filename}.")

def choose_excel_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if file_path:
        with open("path.txt", "w") as f:
            f.write(file_path)
            print("Path saved to path.txt:", file_path)
    else:
        print("No file selected. Creating a blank 'path.txt' file.")
        with open("path.txt", "w") as f:
            f.write("")

if __name__ == "__main__":
    check_and_install_modules()
    download_file("https://raw.githubusercontent.com/Shreyas100100/SEIPL-Project/main/main.py", "main.py")
    download_file("https://raw.githubusercontent.com/Shreyas100100/SEIPL-Project/main/run_main.py", "run_main.py")
    choose_excel_path()
