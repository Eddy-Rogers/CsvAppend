from multiprocessing.context import Process

import openpyxl
import os
import argparse

from tkinter import Tk, Label, filedialog, Button, Frame

inputvalidation = False
outputvalidation = False

global inputpath
global outputpath
inputpath = ""
outputpath = ""

global inputvalid
global outputvalid
inputvalid = False
outputvalid = False

#def validate():
#    return inputvalidation & outputvalidation

#def validcheck():
#    if validate():
#        ProcessFiles(inputpath, outputpath)

def onclickinput():
    global inputpath
    global inputvalid
    inputpath = filedialog.askdirectory()
    inputpathlabel.config(text = "Selected Input Path: " + inputpath)
    inputvalid = True
    if (validate()):
        execbutton.config(bg="green")

def onclickoutput():
    global outputpath
    global outputvalid
    outputpath = filedialog.askdirectory()
    outputpathlabel.config(text = "Selected Output Path: " + outputpath)
    outputvalid = True
    if(validate()):
        execbutton.config(bg = "green")

def validate():
    global inputvalid
    global outputvalid
    return inputvalid == True and outputvalid == True

def execute():
    global inputpath
    global outputpath
    global inputvalid
    global outputvalid
    if(validate()):
        ProcessFiles(inputpath, outputpath)

def gui():
    global inputpath
    global outputpath
    # Simply set the theme
    root.tk.call("source", "azure.tcl")
    root.tk.call("set_theme", "dark")
    root.mainloop()

def num_to_excel_col(n):
    if n < 1:
        raise ValueError("Number must be positive")
    result = ""
    while True:
        if n > 26:
            n, r = divmod(n - 1, 26)
            result = chr(r + ord('A')) + result
        else:
            return chr(n + ord('A') - 1) + result

def ProcessFiles(inputPath, outputPath):
    # Adding some test comments
    directory = os.fsencode(inputPath)
    wb = openpyxl.Workbook()
    ws = wb.active
    rowcounter = 1
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        read = openpyxl.load_workbook(inputPath + "/" + filename)
        dataframe1 = read.active
        counter = rowcounter
        columncounter = 1
        # Iterate the loop to read the cell values
        for row in range(0, dataframe1.max_row):
            for col in dataframe1.iter_cols(1, dataframe1.max_column):
                ws[num_to_excel_col(columncounter) + str(counter)] = col[row].value
                columncounter+=1
            counter += 1
        rowcounter += 1
    wb.save(outputPath + "/combined.xlsx")

root = Tk()
inputlabel = Label(root, text="1. Please select a directory with CSV files.")
inputlabel.pack(padx=10, pady=10)
inputbutton = Button(root, text="Input Path", command=onclickinput)
inputbutton.pack(padx=10, pady=10)
inputpathlabel = Label(root, text="No Input Path Selected")
inputpathlabel.pack(padx=10, pady=10)
frame1 = Frame(root, bg="black", bd=2, relief="solid")
frame1.pack(fill="both", expand=True)
outputlabel = Label(root, text="2. Please select a directory for the aggregate file.")
outputlabel.pack(padx=10, pady=10)
outputbutton = Button(root, text="Output Path", command=onclickoutput)
outputbutton.pack(padx=10, pady=10)
outputpathlabel = Label(root, text="No Output Path Selected")
outputpathlabel.pack(padx=10, pady=10)
frame2 = Frame(root, bg="black", bd=2, relief="solid")
frame2.pack(fill="both", expand=True)
execlabel = Label(root, text="3. Press the button to combine files.")
execlabel.pack(padx=10, pady=10)
execbutton = Button(root, bg="red", text="Execute", command=execute)
execbutton.pack(padx=10, pady=10, expand=True)

def main():
    gui()

if __name__ == "__main__":
    main()