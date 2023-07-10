## Belarcs.py

### Problem Statement:

The code provided is a Python script that allows users to search for specific information within HTML files and generate an Excel output file containing the extracted data. The script utilizes the Tkinter library to create a simple graphical user interface (GUI) for the application.

The main functionality of the script includes:
- Selecting one or more HTML files to search within.
- Entering a search text to look for specific information within the HTML files.
- Choosing an output folder where the generated Excel file will be saved.
- Running the search and file generation process.

The script uses the BeautifulSoup library to parse the HTML content of the selected files and extract relevant information. The extracted data is then written to an Excel file using the openpyxl library. Make sure the required Python libraries (`tkinter`, `openpyxl`, `bs4`) are installed before running the script.

### Code:

```python
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
```

### Procedure:

1. Launch the application by running the script.
2. The GUI window titled "Search Files and Generate Output" will appear.
3. Click the "Browse" button next to the "File(s):" label to select one or more HTML files to search within. You can select multiple files by holding down the Ctrl key while clicking.
4. The selected file paths will be displayed in the corresponding entry field.
5. Enter the desired search text in the "Search Text:" entry field. This text will be used to find specific information within the HTML files.
6. Click the "Choose" button next to the "Output Folder Path:" label to select the folder where the generated Excel file will be saved.
7. The selected output folder path will be displayed in the corresponding entry field.
8. Click the "Run" button to start the search and file generation process.
9. If any of the required fields (file paths, search text, or output folder path) are empty, an error message will be displayed. Make sure to provide values for all required fields.
10. If one or more of the selected files are invalid (not existing or inaccessible), an error message will be displayed.
11. Once the search and file generation process is completed, a success message will be displayed.
12. The generated Excel file will be saved in the chosen output folder with the name "output.xlsx".

### Outcome:
```
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC017_IT_MAHESHMARADANA_VSPITP_1F_D76.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC017_IT_SREEYA_VSPITP_1F_D76.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC017_IT_SUDHEER_VSPITP_1F_D76.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC017_IT_VINODH_VSPITP_1F_D76.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC022_IT_ANIL_VSPITP_1F_D87.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC022_IT_SUNIL_VSPITP_1F_D87.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC033_IT_SUBBARAJU_VSPITP_1F_D74.html
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC046_IT_APPALANAIDU_VSPITP_1F_D73.html
Excel file saved successfully!
```

After parsing the HTML file and extracting the relevant information, the script created an Excel file named "output.xlsx". The Excel file was saved successfully, indicating that the search and file generation process completed without errors.
