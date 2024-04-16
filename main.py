from startup import Book
import openpyxl #nessisary for formatting outfile
from pathlib import Path #to make pathing easier maybe

#used by getlist to append hyperlink
def placeHlink(inPath,datalist):
    hStr = inPath / "cData" 
    hStr = hStr.resolve()
    #append list with hyperlink string for folder path
    datalist.append('=HYPERLINK("{}")'.format(hStr))
    #return list
    return datalist

def getList(inPath):
    dPath = inPath / "data.txt"
    data = []
    #make list of data.txt shit
    with dPath.open() as file:
        for line in file:
            key, value = line.strip().split(':')
            data.append(value)
    #paceHlink and return
    return placeHlink(inPath,data)
    
 
def writeRow(datalist):
    global book
    global currentRow
    for col_idx, cellValue in enumerate(datalist, start=1):
        #shortcut Current Cell
        currCell= book.s.cell(row=currentRow, column=col_idx)
        #store cell value in current cell
        currCell.value=cellValue 
        #if its the last entry in list set cell style to hyperlink
        if(col_idx==len(datalist)):
            currCell.style="Hyperlink"
    currentRow+=1
    return(datalist)

def combDataFolder():
    #make list of folders in datafolder
    dataDir = Path("./") / "dataFolder"
    entries = [Dir for Dir in dataDir.iterdir() if Dir.is_dir()]

    #now for each folder containing client data:
    for entry in entries:
        datalist = getList(Path(entry))
        writeRow(datalist)

if __name__ == "__main__":
    #start Book up
    book = Book()
    book.main()
    #set row for outfile
    currentRow =2
    combDataFolder()
    book.save()

"""
Depreciated

# def read_key_value_pairs(file_path='data.txt'):
#     # Initialize an empty dictionary to store key-value pairs
#     data = {}
#
#     # Open the file for reading
#     with open(file_path, 'r') as file:
#         # Read each line in the file
#         for line in file:
#             # Split each line into key and value using the ':' delimiter
#             key, value = line.strip().split(':')
#             # Store the key-value pair in the dictionary
#            data[key] = value


"""
