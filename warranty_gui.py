import warranty_lookup as wnty
import tkinter as tk
from tkinter import filedialog as fd


root = tk.Tk()
root.title("Warranty Lookup")
#main application menu
menu = tk.Menu(root)
root.config(menu=menu)

file_menu = tk.Menu(menu)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=open_file)


def open_file():
    csv_file = fd.askopenfilename(title="Open CSV File", filetypes=("CSV Files", "*.csv"))
    wnty = wnty.process_csv(csv_file)
    print("CSV file processed successfully.")

def download_warranty_info():
    pass # Implement functionality to download warranty information from the API
#"C:\Coding Projects\Warranty Project"