# Belarcs.py

### Problem Statement

The code provided is a Python script that allows users to search for specific information within HTML files and generate an Excel output file containing the extracted data. The script utilizes the Tkinter library to create a simple graphical user interface (GUI) for the application.

The main functionality of the script includes:
- Selecting one or more HTML files to search within.
- Entering a search text to look for specific information within the HTML files.
- Choosing an output folder where the generated Excel file will be saved.
- Running the search and file generation process.

The script uses the BeautifulSoup library to parse the HTML content of the selected files and extract relevant information. The extracted data is then written to an Excel file using the openpyxl library.

### Process

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

Note: Make sure the required Python libraries (`tkinter`, `openpyxl`, `bs4`) are installed before running the script.
