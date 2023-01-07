import tkinter as tk
from tkinter import BOTH, END, LEFT
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import pandas as pd
import datetime
from pathlib import Path
import os

window_height = 500
window_width = 500

root = tk.Tk()
root.title('Merger App Test')
root.resizable(False, False)
root.geometry('500x500')

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))

root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))


lblTitle = ttk.Label(text="Excel File Merger", font=("Roboto",25), justify='center')
lblTitle.grid(row=0, column=1)


#Button One Section
lblButtonOne = ttk.Label(text="First file: ", font=("Roboto",11), justify='left')
lblButtonOne.grid(row=1, column=0,pady=40)

txtFileOne = ttk.Entry(font=("Roboto",11), width=25)
txtFileOne.grid(row=1, column=1,pady=40)
txtFileOne.configure(state="readonly")

def select_file1():
    filetypes = (
        ("Excel files", ".xlsx .xls"),
        ("All files", ".")
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    txtFileOne.configure(state="default")
    txtFileOne.delete(0,END)
    txtFileOne.insert(0,filename)
    txtFileOne.configure(state="readonly")

btnFileOne = ttk.Button(text="Select File", command=select_file1)
btnFileOne.grid(row=1,column=2,pady=40)
#Button One Section Ends

#Button Two Section
lblButtonTwo = ttk.Label(text="Second file: ", font=("Roboto",11))
lblButtonTwo.grid(row=2, column=0,pady=40)

txtFileTwo = ttk.Entry(font=("Roboto",11), width=25)
txtFileTwo.grid(row=2, column=1,pady=40)
txtFileTwo.configure(state="readonly")

def select_file2():
    filetypes = (
        ("Excel files", ".xlsx .xls"),
        ("All files", ".")
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    txtFileTwo.configure(state="default")
    txtFileTwo.delete(0,END)
    txtFileTwo.insert(0,filename)
    txtFileTwo.configure(state="readonly")

btnFileTwo = ttk.Button(text="Select File", command=select_file2)
btnFileTwo.grid(row=2,column=2,pady=40)
#Button Two Section Ends


#merge section
lblHeader = ttk.Label(text="Header to Merge on", font=("Roboto",11))
lblHeader.grid(row=3, column=0)

txtHeader = ttk.Entry(font=("Roboto",11), width=25)
txtHeader.grid(row=3, column=1, columnspan=1)
def merge_sheets(header):
    if len(header) > 0:
        sheet1 = pd.read_excel(str(txtFileOne.get()),sheet_name=0)
        sheet2 = pd.read_excel(str(txtFileTwo.get()),sheet_name=0)
        
        if header in sheet1 and header in sheet2:
            df = sheet1.merge(sheet2, on = header, how='outer')

            df.loc[df[header].duplicated(), header] = pd.NA

           
            today = datetime.datetime.now()
            today = today.strftime("%Y-%m-%d_%H-%M")

            path = f'{Path.cwd()}/VHT Merged_{today}.xlsx'
            
             #export new dataframe to excel
            df.to_excel(f'VHT Merged_{today}.xlsx')

            showinfo(
                title='New Sheet Made',
                message=f"File can be found at:\n{Path.cwd()}/VHT Merged_{today}.xlsx"
            )
            os.startfile(path)

            txtFileOne.delete(0,END)
            txtFileTwo.delete(0,END)
            txtHeader.delete(0,END)
        else:
            showinfo(
            title='HEADER REQUIRED',
            message='Header needs to exsist within both sheets.'
        )
    else:
        showinfo(
            title='HEADER REQUIRED',
            message='Header is required to merge the sheets.'
        )


btnMerge = ttk.Button(text="Merge Files", command=lambda: merge_sheets(str(txtHeader.get())))
btnMerge.grid(row=4,column=1,pady=40,padx=20)
#merge section ends



root.mainloop()
