import os

keyword = input("enter a keyword: \n")

directory = os.fsencode("/home/karpi/Documents/excelsample/")
    
for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        # print(os.path.join(directory, filename))
        wb = openpyxl.load_workbook(filename)
        sheet = wb.get_sheet_by_name('Sheet1')
        for i in range(1,10):
            for j in range(1,10):
                if sheet.cell(row = i, column = j).value == keyword:
                print("keywoard is in document")
                
        continue
    else:
        continue