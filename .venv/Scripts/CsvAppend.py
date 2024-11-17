import openpyxl
import os
import argparse

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
    # another test
    # third test
    directory = os.fsencode("./" + inputPath)
    wb = openpyxl.Workbook()
    ws = wb.active
    rowcounter = 1
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        read = openpyxl.load_workbook("./" + inputPath + "/" + filename)
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
    wb.save("./" + outputPath + "/sample.xlsx")

def main():
    parser = argparse.ArgumentParser(description="This script appends excel files to eachother.")

    # Add a positional argument
    parser.add_argument("InputPath", help="The name of the folder to pull files from.")

    # Add a positional argument
    parser.add_argument("OutputPath", help="The name of the folders to put files into.")

    # Parse the arguments
    args = parser.parse_args()

    ProcessFiles(vars(args)["InputPath"], vars(args)["OutputPath"])

if __name__ == "__main__":
    main()