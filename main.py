import os
import openpyxl

keyword = input("enter a keyword: \n")
excel_path = "/home/karpi/Documents/excelsample/"
txt_path = "/home/karpi/Documents/"

#make a directory object which contains the directory to use
directory = os.fsencode(excel_path)

#change directory to .txt path so we can create a file there
os.chdir(txt_path)

#opens and rewrites, or creates a 
SearchResults = open("SearchResult.txt","w")

#change directory path to the excel spreadsheets so we can read them
os.chdir(excel_path)


#go through files in a directory
for file in os.listdir(directory):
    filename = os.fsdecode(file)

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        
        #read excel file
        wb = openpyxl.load_workbook(filename)

        #read sheets in order
        for sheet in wb:
            #2d for loop to read from cells 
            for i in range(1,10):
                for j in range(1,10):
                    #read cell value and convert them to string
                    #so we can use lower method on them
                    cellValue = str(sheet.cell(row = i, column = j).value)
                    if keyword.lower() == cellValue.lower():
                        #write the file name which contains the keyword into a .txt file
                        SearchResults.write(filename)
                        SearchResults.write("\n")
                        break
                else:
                    continue
                break
                
        continue
    else:
        continue

SearchResults.close()