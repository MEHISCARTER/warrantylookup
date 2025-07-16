import warranty_lookup as wnty
import tkinter as tk
from tkinter import filedialog as fd, ttk, messagebox
from tkinter import StringVar
import csv

csv_file = None
tree = None
table_frame = None
filter_var = None
filter_col_var = None
match_type_var = "Contains"
all_rows = []

"""
Download a csv file after inputting it into the "Enter file" button

Add filter for columns - done
Finish help menu - done
Finish about menu - done
Add download filtered file
Add editing function
Add other QOL changes I can think of
"""

def open_file():
    global csv_file

    csv_file = fd.askopenfilename(title="Open CSV File", filetypes=(("CSV Files", "*.csv"),))
    if csv_file:    
        csv_file = wnty.process_csv(csv_file)
        show_csv_tree(csv_file)

def show_csv_tree(file_path):
    global tree, all_rows, filter_var, filter_col_var, table_frame, match_type_var

    for widget in table_frame.winfo_children():
        widget.destroy()
    
    with open(file_path, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    if not rows:
        return
    
    columns = rows[0]
    all_rows = rows

    filter_frame = tk.Frame(table_frame)
    filter_frame.pack(fill='x', padx=5, pady=5)

    filter_col_var = StringVar(value = columns[0])
    filter_var = StringVar()
    match_type_var = StringVar(value='Contains')
    
    match_type_menu = ttk.Combobox(filter_frame, textvariable=match_type_var, values=['Contains', 'Exact'], state='readonly', width=10)
    match_type_menu.pack(side='left', padx=5)

    tk.Label(filter_frame, text= "Filter by:").pack(side='left')
    col_menu = ttk.Combobox(filter_frame, textvariable=filter_col_var, values=columns, state= "readonly", width=15)
    col_menu.pack(side='left', padx=5)
    tk.Entry(filter_frame, textvariable=filter_var, width=20).pack(side='left', padx=5)
    tk.Button(filter_frame, text="Apply Filter", command=apply_filter).pack(side='left', padx=5)
    tk.Button(filter_frame, text="Clear Filter", command=clear_filter).pack(side='left', padx=5)

    #Table
    tree = ttk.Treeview(table_frame, columns=columns, show='headings')

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')

    for row in rows[1:]:
        tree.insert('', tk.END, values=row)
    tree.pack(fill='both', expand=True)

def apply_filter():
    col = filter_col_var.get()
    val = filter_var.get().lower()

    match_type = match_type_var.get()

    columns = all_rows[0]
    col_idx = columns.index(col)
    tree.delete(*tree.get_children())
    for row in all_rows[1:]:
        cell = row[col_idx].lower()
        if match_type == "Contains":
            if val in cell:
                tree.insert('', tk.END, values=row)
        elif match_type == "Exact":
            if cell == val:
                tree.insert('', tk.END, values=row)

def clear_filter():
    filter_var.set('')
    apply_filter()

def main(): 
    global table_frame
    root = tk.Tk()
    root.title("Warranty Lookup")
    root.state('zoomed') #Maximize window automatically for Windows

    #main application menu
    menu = tk.Menu(root)
    root.config(menu=menu)

    #help menu
    help_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="How to use", command=show_help)
    help_menu.add_command(label="About", command=show_about)

    #file menu
    file_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open File for Warranty Lookup", command=open_file)
    file_menu.add_command(label="Download Updated File", command=download_file)

    #Frame for table
    table_frame = tk.Frame(root)
    table_frame.pack(fill='both', expand=True)

    root.mainloop()

def download_file():
    if csv_file:
        save_path = fd.asksaveasfilename(
            defaultextension=".csv",
            filetypes=(("CSV Files", "*.csv"),),
            title="Save CSV File As"
        )
        if save_path:
            with open(csv_file, 'rb') as src, open(save_path, 'wb') as dest:
                dest.write(src.read())
                messagebox.showinfo("CSV File Downloaded!", f"File saved to:\n{save_path}")
        else:
            messagebox.showinfo("Cancelled", "Download cancelled!")
    else:
        messagebox.showwarning("No File", "Open a CSV file first to download it!")



def show_help():
    tk.messagebox.showinfo("How to Use", "1: Create Excel or Sheets file with list of serial numbers you want to look up." \
    "\n2. Download those files as CSV files." \
    "\n3. Open the CSV file using File -> Open File for Warranty Lookup." \
    "\n4. The table will populate on the main screen with the serial numbers and their warranty status." \
    "\n5. Download the updated file if it is to your liking with File -> Download Updated File")
def show_about():
    tk.messagebox.showinfo("About", "This application was created using Python and Tkinter."
    "\nIts purpose is to look up Lenovo warranty information from a CSV file containing the serial numbers of devices."
    "\nThere are some limitations, such as some Lenovo devices may have a serial number linked with multiple devices."
    "\nThis application cannot find those devices correctly, so you must go to Lenovo's website to find them."
    "\nThis was created by a student from Grand Valley State University and has no affiliation with Lenovo.")

if __name__ == "__main__":
    main()