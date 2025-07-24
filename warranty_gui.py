import warranty_lookup as wnty
import tkinter as tk
from tkinter import filedialog as fd, ttk, messagebox, StringVar, font as tkFont
import csv, os

csv_file = None
tree = None
table_frame = None
copy_frame = None
filter_var = None
field_vars = None
filter_col_var = None
match_type_var = "Contains"
all_rows = []
fields = []

"""
Download a csv file after inputting it into the "Enter file" button

Add filter for columns - done
Finish help menu - done
Finish about menu - done
Add download filtered file - done
Add copy function for tags/rows - done
Add progress - done
Add vertical scroll bar - done
Add editing function - maybe?
Add other QOL changes I can think of
"""

def main(): 
    global table_frame, placeholder_label
    root = tk.Tk()
    root.title("Warranty Lookup")
    root.state('zoomed') #Maximize window automatically for Windows

    #main application menu
    menu = tk.Menu(root)
    root.config(menu=menu)

    #Frame for table
    table_frame = tk.Frame(root)
    table_frame.pack(fill='both', expand=True)

    placeholder_label = tk.Label(table_frame, text="Select a CSV file to begin!", anchor='w', font=('Arial', 30), fg='black')
    placeholder_label.pack(expand=True)

    #help menu
    help_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="How to use", command=show_help)
    help_menu.add_command(label="About", command=show_about)
    #help_menu.add_command(label="Lenovo Name Scheme", command=show_lenovo_scheme)

    #file menu
    file_menu = tk.Menu(menu, tearoff=0)
    menu.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open File for Warranty Lookup", command=open_file)
    file_menu.add_command(label="Download Updated File", command=download_file)
    file_menu.add_command(label="Download Filtered File", command=download_filter)

    root.mainloop()

def open_file():
    global csv_file

    csv_file = fd.askopenfilename(title="Open CSV File", filetypes=(("CSV Files", "*.csv"),))
    if csv_file:
        file_name = os.path.basename(csv_file) 
        
        for widget in table_frame.winfo_children():
            widget.destroy()

        progress_label = tk.Label(table_frame, text=f"Processing {file_name}...", anchor='w', font=('Arial', 14), fg='black')
        progress_label.pack(fill='x', padx=10, pady=10)
        progress_label.update_idletasks()

        def update_progress(current, total):
            progress_label.config(text=f"Processing {file_name}... {current}/{total} - {(current/total)*100:.2f}%")
            progress_label.update_idletasks()

        csv_file = wnty.process_csv(csv_file, progress_callback=update_progress)
        progress_label.config(text=f"Finished processing {file_name}")
        show_csv_tree(csv_file)

def show_csv_tree(file_path):
    global tree, all_rows, filter_var, filter_col_var, table_frame, match_type_var, field_vars, asset_entry, fields, copy_frame

    for widget in table_frame.winfo_children():
        widget.destroy()
    
    with open(file_path, newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    if not rows:
        return
    
    columns = rows[0]
    all_rows = rows

    #filter frame
    filter_frame = tk.Frame(table_frame)
    filter_frame.pack(fill='x', padx=5, pady=5)

    #copy frame
    copy_frame = tk.Frame(table_frame)
    copy_frame.pack(fill='x', padx=10)

    #filter variables
    filter_col_var = StringVar(value = columns[0])
    filter_var = StringVar()
    match_type_var = StringVar(value='Contains')
    
    #contains or exact match search filter
    match_type_menu = ttk.Combobox(filter_frame, textvariable=match_type_var, values=['Contains', 'Exact'], state='readonly', width=10)
    match_type_menu.pack(side='left', padx=5)

    #buttons and labels
    tk.Label(filter_frame, text= "Filter by:").pack(side='left')
    col_menu = ttk.Combobox(filter_frame, textvariable=filter_col_var, values=columns, state= "readonly", width=15)
    col_menu.pack(side='left', padx=5)
    tk.Entry(filter_frame, textvariable=filter_var, width=20).pack(side='left', padx=5)
    tk.Button(filter_frame, text="Apply Filter", command=lambda: [apply_filter(), update_asset_tag_dropdown()]).pack(side='left', padx=5)
    tk.Button(filter_frame, text="Clear Filter", command=clear_filter).pack(side='left', padx=5)
    #tk.Button(copy_frame, text="Copy Asset Info", command=copy_asset_info).pack(side='left', padx=5)

    #Table
    style = ttk.Style()
    style.configure("Treeview", font=('Consolas', 8, 'normal'))  # Change row font
    style.configure("Treeview.Heading", font=('Consolas', 10, 'bold'))
    tree = ttk.Treeview(table_frame, columns=columns, show='headings')

    #horizontal scrollbar
    xscroll = tk.Scrollbar(table_frame, orient='horizontal', command=tree.xview, width=20, troughcolor='black')
    tree.configure(xscrollcommand=xscroll.set)
    xscroll.pack(side='bottom', fill='x')

    #vertical scrollbar
    yscroll = tk.Scrollbar(table_frame, orient='vertical', command=tree.yview, width=20, troughcolor='black')
    tree.configure(yscrollcommand=yscroll.set)
    yscroll.pack(side='right', fill='y')

    font = tkFont.Font(font=('Consolas', 8, 'normal'))
    for col in columns:
        tree.heading(col, text=col)
        max_width = font.measure(col)
        #print(f'max width initial: {max_width}')
        for row in rows[1:]:
            cell_text = str(row[columns.index(col)])
            cell_width = font.measure(cell_text)
            #print(f'cell text: {cell_text}, cell width: {cell_width}')
            if cell_width > max_width:
                max_width = cell_width
    tree.column(col, width=max_width, anchor='center')

    for row in rows[1:]:
        tree.insert('', tk.END, values=row)

    tree.pack(fill='both', expand=True)

    #asset tag dropdown
    def get_visible_assets():
        idx = columns.index("Asset Tag")
        return [tree.item(item)["values"][idx] for item in tree.get_children()]

    asset_tag_var = StringVar()
    asset_tag_dropdown = ttk.Combobox(copy_frame, textvariable=asset_tag_var, width=20, state='readonly')
    asset_tag_dropdown.pack(side='left', padx=5)

    def update_asset_tag_dropdown():
        tags = get_visible_assets()
        asset_tag_dropdown['values'] = tags
        if tags:
            asset_tag_var.set(tags[0])
        else:
            asset_tag_var.set('')
    
    update_asset_tag_dropdown()

    #Check boxes
    fields = ["Asset Tag", "Serial Number", "Start/Purchase Date", "Warranty End", "Product Model", "NetID/Name"]
    field_vars = {field: tk.BooleanVar(value=True) for field in fields}
    
    #copy button for field
    def copy_single_field(field):
        asset_tag = asset_tag_var.get().strip()
        if not asset_tag:
            messagebox.showwarning("Input Error", "Please select an Asset Tag to copy.")
            return
        try:
            asset_idx = columns.index("Asset Tag")
            field_idx = columns.index(field)
        except ValueError:
            messagebox.showerror("Error", f"{field} column not found in CSV.")
            return
        for row in all_rows[1:]:
            if row[asset_idx].strip() == asset_tag:
                value = row[field_idx]
                table_frame.clipboard_clear()
                table_frame.clipboard_append(value)
                messagebox.showinfo("Copied", f"Copied {field}: {value}")
                return
        messagebox.showwarning("Not Found", f"No entry found for Asset Tag: {asset_tag}")
    
    for field in fields:
        tk.Button(copy_frame, text=f"Copy {field}", command=lambda f=field: copy_single_field(f)).pack(side='left', padx=2)


#     tk.Label(copy_frame, text="Copy using Asset Tag:").pack(side='left', padx=5)
#     asset_entry = tk.Entry(copy_frame, width=20)
#     asset_entry.pack(side='left', padx=5)

#     #Checkboxes
#     fields = ["Asset Tag", "Serial Number", "Start/Purchase Date", "Warranty End", "Product Model", "NetID/Name"]
#     field_vars = {field: tk.BooleanVar(value=True) for field in fields}
#     for field in fields: 
#         tk.Checkbutton(copy_frame, text=field, variable=field_vars[field]).pack(side='left', padx=2)

# def copy_asset_info():
#     asset_tag = asset_entry.get().strip()
#     if not asset_tag:
#         messagebox.showwarning("Input Error", "Please enter an Asset Tag to copy.")
#         return
    
#     columns = all_rows[0]
#     try:
#         asset_idx = columns.index("Asset Tag")
#     except ValueError:
#         messagebox.showerror("Error", "Asset Tag column not found in CSV.")
#         return
#     found = False

#     for row in all_rows[1:]:
#         if row[asset_idx].strip() == asset_tag:
#             table_frame.clipboard_clear()
#             copied = []
#             for field in fields:
#                 if field_vars[field].get():
#                     try:
#                         idx = columns.index(field)
#                         value = row[idx]
#                         table_frame.clipboard_append(value)
#                         copied.append(f"{field}: {value}")
#                     except ValueError:
#                         pass
#             if copied:
#                 messagebox.showinfo("Copied", f"Copied: {', '.join(copied)}")
#             else:
#                 messagebox.showwarning("No Fields Selected", "No fields were selected to copy.")
#             found = True
#             break
#     if not found:
#         messagebox.showwarning("Not Found", f"No entry found for Asset Tag: {asset_tag}")

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

def download_filter():
    if tree is None:
        messagebox.showwarning("No Filtered Data", "Apply Filter first to download a filtered file!")
        return
    save_path = fd.asksaveasfilename(
        defaultextension=".csv",
        filetypes=(("CSV Files", "*.csv"),),
        title="Save Filtered File As"
    )

    if not save_path:
        messagebox.showinfo("Cancelled", "Download cancelled!")
        return
    
    columns = tree["columns"]

    data = [tree.item(item)["values"] for item in tree.get_children()]

    with open(save_path, "w", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        writer.writerows(data)
    
    messagebox.showinfo("Filtered File Downloaded!", f"File saved to:\n{save_path}")

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

def show_lenovo_scheme():
    pass

if __name__ == "__main__":
    main()