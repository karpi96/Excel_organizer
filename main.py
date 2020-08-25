import os
import openpyxl
from tkinter import *
from tkinter import filedialog

#makes a tkinter window and calls it root
root = Tk()

#make a entry widgets 50 pixels wide
entryText = Entry(root,width = 50)
entryText.insert(0,"Here goes the path for the search results.")
entryExcel = Entry(root,width = 50)
entryExcel.insert(0,"Here goes the path for the excel spreadsheets.")
entryKeyword = Entry(root,width = 50)
entryKeyword.insert(0,"Enter Keyword")

#put the entry widgets in a grid
entryText.grid(row = 0, column = 0)
entryExcel.grid(row = 1, column = 0)
entryKeyword.grid(row = 2, column = 0)


#this function opens the file dialog box, asks for a directory
#and then writes the path into the .txt entry
def askFolderText():
    root.filename = filedialog.askdirectory()
    entryText.delete(0,END)
    entryText.insert(0,root.filename)

#this function opens the file dialog box, asks for a directory
#and then writes the path into the excel entry
def askFolderExcel():
    root.filename = filedialog.askdirectory()
    entryExcel.delete(0,END)
    entryExcel.insert(0,root.filename)

#this function generates a text file where the file names which contain the selected keyword are written in
def generate():

    
    excel_path = entryExcel.get()
    txt_path = entryText.get()
    keyword = entryKeyword.get()
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

#button that executes the askFolderText function 
myButton = Button(root, text="Txt location select", command=askFolderText,width = 30)
myButton.grid(row = 0, column = 1)

#button that executes the askFolderExcel function 
myButton = Button(root, text="Excel location select", command=askFolderExcel, width = 30)
myButton.grid(row = 1, column = 1)

#button that generates the file
myButton = Button(root, text="Generate file", command=generate, width = 30)
myButton.grid(row = 2, column = 1)

root.mainloop()
