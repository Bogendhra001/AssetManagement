import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import re
import openpyxl
from bs4 import BeautifulSoup

def search_files(file_paths, search_text, output_folder_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = 'Output'

    # Set column headers
    worksheet.append(['SystemName', 'Department', 'Employee Name', 'Branch', 'Floor', 'Port', 'System Model',
                      'Processor', 'Main Circuit Board', 'Drives','Memory Modules',  'Display'])

    row = 2  # Start from row 2 for data

    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8') as file_content:
            print(f"File: {file_path}")

            # Load the HTML into BeautifulSoup
            soup = BeautifulSoup(file_content, 'html.parser')

            # System Model
            desired_caption = 'System Model'
            table = soup.find('caption', string=re.compile(desired_caption)).find_parent('table') if soup.find('caption', string=re.compile(desired_caption)) else None

            system_model = ''
            if table:
                html_content = table.find('td').decode_contents().strip()
                lines = html_content.split('<br>')
                if len(lines) >= 2:
                    system_model = lines[0].strip() + '\n' + lines[1].strip()
                else:
                    system_model = BeautifulSoup(html_content, 'html.parser').get_text().strip()

            # Processor
           

            div2 = soup.find_all("div",{'class':"reportSection rsLeft"})
            # print(div2)
            if len(div2) >= 2:
                second_div = (div2[1])
                # second_div2= (div2[2])
            # print(second_div)
            # processor = second_div.find('td').get_text(strip=True)
            processor = second_div.find('td').get_text(strip=True)
            # print(processor)
            # print(td_text)
            # processor=td_text
            # div3 = soup.find_all("div",{'class':"reportSection rsRight"})
            # print(len(div3))
            # print(div3[1])
            # if len(div3) >= 2:
            #     second_div1 = (div3[4])
            # # print(second_div1)
            # processor = second_div1.find('td').get_text(strip=True)


            # Main Circuit Board
            # desired_caption3 = 'Main Circuit Board b'
            # table3 = soup.find('caption', string=re.compile(desired_caption3)).find_parent('table') if soup.find('caption', string=re.compile(desired_caption3)) else None

            # main_circuit_board = table3.find('td').contents[0].strip() if table3 else ''
            div2 = soup.find_all("div",{'class':"reportSection rsRight"})
            # print(div2)
            # print(div2[1])
            a=div2[1]
            td_text = a.find('td').get_text(strip=True)
            # print(td_text)
            # print(type(td_text))
            main_circuit_board=td_text
            

            # Drives
            desired_caption4 = 'Drives'
            table4 = soup.find('caption', string=re.compile(desired_caption4)).find_parent('table') if soup.find('caption', string=re.compile(desired_caption4)) else None

            drives = table4.find('td').contents[0].strip() if table4 else ''

            # Memory Modules
            # desired_caption5 = 'Memory Modules c,d'
            # table5 = soup.find('caption', string=re.compile(desired_caption5)).find_parent('table') if soup.find('caption', string=re.compile(desired_caption5)) else None

            # memory_modules = table5.find('td').contents[0].strip() if table5 else ''

            div2 = soup.find_all("div",{'class':"reportSection rsRight"})
            # print(div2)
            # print(div2[1])
            a=div2[2]
            # print(a)
            td_text = a.find('td').get_text(strip=True)
            # print(td_text)
            memory_modules=td_text
            # Display
            desired_caption6 = 'Display'
            table6 = soup.find('caption', string=re.compile(desired_caption6)).find_parent('table') if soup.find('caption', string=re.compile(desired_caption6)) else None

            display = table6.find('td').decode_contents().split('<br>')[0].strip() if table6 else ''

            SystemName = os.path.basename(file_path).split('_')[0]
            Dept = os.path.basename(file_path).split('_')[1]
            ename = os.path.basename(file_path).split('_')[2]
            branch = os.path.basename(file_path).split('_')[3]
            sym_floor = os.path.basename(file_path).split('_')[4]
            port = os.path.basename(file_path).split('_')[5].split('.')[0]

            # Add data to Excel worksheet
            worksheet.append([SystemName, Dept, ename, branch, sym_floor, port, system_model, processor,
                              main_circuit_board, drives,  memory_modules,display])

    # Save the Excel file
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


# Create the main window
window = tk.Tk()
window.title("Search Files and Generate Output")
window.geometry("400x250")

# Create and position the widgets
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

# Run the main event loop
window.mainloop()