import os
import openpyxl

keyword = input("enter a keyword: \n")
os.chdir("/home/karpi/Documents/excelsample/")


directory = os.fsencode("/home/karpi/Documents/excelsample/")

file_name = open("SearchResult.txt","w")

#go through files in a directory
for file in os.listdir(directory):
    filename = os.fsdecode(file)

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        
        #read excel file
        wb = openpyxl.load_workbook(filename)
        sheet = wb.get_sheet_by_name('Sheet1')
        for i in range(1,10):
            for j in range(1,10):
                if sheet.cell(row = i, column = j).value == keyword:
                    file_name.write(filename)
                    break
            else:
                continue
            break
                
        continue
    else:
        continue