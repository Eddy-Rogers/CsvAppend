from multiprocessing.context import Process

import openpyxl
import os
import argparse

from tkinter import Tk, Label, filedialog, Button

inputvalidation = False
outputvalidation = False

global inputpath
global outputpath




#def validate():
#    return inputvalidation & outputvalidation

#def validcheck():
#    if validate():
#        ProcessFiles(inputpath, outputpath)


def gui():
    def onclickinput():
        global inputpath
        inputpath = filedialog.askdirectory()
    def onclickoutput():
        global outputpath
        outputpath = filedialog.askdirectory()
    def execute():
        global inputpath
        global outputpath
        ProcessFiles(inputpath, outputpath)
    root = Tk()
    inputbutton = Button(root, text="Input Path", command=onclickinput)
    inputbutton.pack()
    outputbutton = Button(root, text="Output Path", command=onclickoutput)
    outputbutton.pack()
    execbutton = Button(root, text="Execute", command=execute)
    execbutton.pack()
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

def main():
    gui()

if __name__ == "__main__":
    main()