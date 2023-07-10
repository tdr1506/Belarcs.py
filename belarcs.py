import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import openpyxl
from bs4 import BeautifulSoup

def extract_value(soup, caption_text):
    table = soup.find('caption', string=re.compile(caption_text)).find_parent('table') if soup.find('caption', string=re.compile(caption_text)) else None
    if table:
        return table.find('td').get_text(strip=True)
    return ''

def search_files(file_paths, search_text, output_folder_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = 'Output'
    worksheet.append(['SystemName', 'Department', 'Employee Name', 'Branch', 'Floor', 'Port', 'System Model',
                      'Processor', 'Main Circuit Board', 'Drives','Memory Modules',  'Display'])

    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8') as file_content:
            print(f"File: {file_path}")

            soup = BeautifulSoup(file_content, 'html.parser')

            system_model = extract_value(soup, 'System Model')

            div2 = soup.find_all("div", {'class':"reportSection rsLeft"})
            processor = div2[1].find('td').get_text(strip=True)

            div2 = soup.find_all("div", {'class':"reportSection rsRight"})
            main_circuit_board = div2[1].find('td').get_text(strip=True)
            drives = extract_value(soup, 'Drives')
            memory_modules = div2[2].find('td').get_text(strip=True)
            display = extract_value(soup, 'Display')

            filename = os.path.basename(file_path)
            SystemName, Dept, ename, branch, sym_floor, port = filename.split('_')[:6]

            worksheet.append([SystemName, Dept, ename, branch, sym_floor, port, system_model, processor,
                              main_circuit_board, drives, memory_modules, display])

    output_file_path = os.path.join(output_folder_path, "output.xlsx")
    workbook.save(output_file_path)
    print('Excel file saved successfully!')

def browse_files():
    file_paths = filedialog.askopenfilenames()
    file_paths = [os.path.normpath(file_path) for file_path in file_paths]
    file_entry.delete(0, tk.END)
    file_entry.insert(0, ", ".join(file_paths))

def browse_output_folder():
    output_folder_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_folder_path)

def run_search():
    file_paths_text = file_entry.get()
    search_text = search_entry.get()
    output_folder_path = output_entry.get()

    if not file_paths_text or not search_text or not output_folder_path:
        messagebox.showerror("Error", "Please select file(s), enter search text, and choose an output folder path.")
        return

    file_paths = file_paths_text.split(", ")
    if not all(os.path.isfile(file_path) for file_path in file_paths):
        messagebox.showerror("Error", "One or more selected files are invalid.")
        return

    try:
        search_files(file_paths, search_text, output_folder_path)
        messagebox.showinfo("Success", "Search and file generation completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


window = tk.Tk()
window.title("Search Files and Generate Output")
window.geometry("400x250")

file_label = tk.Label(window, text="File(s):")
file_label.pack()

file_entry = tk.Entry(window, width=50)
file_entry.pack()

browse_button = tk.Button(window, text="Browse", command=browse_files)
browse_button.pack()

search_label = tk.Label(window, text="Search Text:")
search_label.pack()

search_entry = tk.Entry(window, width=50)
search_entry.pack()

output_label = tk.Label(window, text="Output Folder Path:")
output_label.pack()

output_entry = tk.Entry(window, width=50)
output_entry.pack()

output_button = tk.Button(window, text="Choose", command=browse_output_folder)
output_button.pack()

run_button = tk.Button(window, text="Run", command=run_search)
run_button.pack()

window.mainloop()
