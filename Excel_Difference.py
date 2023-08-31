import pandas as pd
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

filename = filedialog.askopenfilename(title="Pick your Excel")

#list number 1 -> Sheet(List1)
#list number 2 -> Sheet(List2)
#list number 3 -> Sheet(Inconsistencies)

class Liste:
    def __init__(self, df) -> None:
        self.df = df
        self.col1 = self.getColumn(self.df, 0) # Column 1
        self.col2 = self.getColumn(self.df, 1) # Column 2
        self.col3 = self.getColumn(self.df, 2) # Column 3
        self.dataArr = [(self.col1).to_list(), (self.col2).to_list(), (self.col3).to_list()]
    
    def getColumn(self, df, columnNr) -> list:
        return df.iloc[:, columnNr]
    
    def getSheet(self) -> None:
        return self.dataArr # ----- [[Column 1],[Column 2],[Column 3]]


#compares two lists at specific column
def compareLists(firstList, secondList, columnNr):
    diff = [] # differences gets stored here
    for item1 in firstList[columnNr]: #iterate through first list
        inList2 = False
        for item2 in secondList[columnNr]: #in every iteration iterate through second list
            if item1 == item2: inList2 = True #compare if any item out of second list is same as the item in first list - if so set flag to true

        if not inList2: diff.append(item1) #if flag is still false then there is no item1 in second list and it gets added to diff array

    #this for loop is the same logic but it compares if any items out of second list are not in first list
    for item2 in secondList[columnNr]:
        inList1 = False
        for item1 in firstList[columnNr]:
            if item2 == item1: inList1 = True
        
        if not inList1: diff.append(item2)

    return diff

# Read the Excel files into DataFrames
xlsx = pd.ExcelFile(filename)

df1 = pd.read_excel(xlsx, "List1")
df2 = pd.read_excel(xlsx, "List2")
df3 = pd.read_excel(xlsx, "Inconsistencies")

list1 = (Liste(df1)).getSheet()
list2 = (Liste(df2)).getSheet()

# differences in Functionvariants ID
col1Diff = compareLists(list1, list2, 0)

# differences in Functionvariants names
col2Diff = compareLists(list1, list2, 1)

# differences in Functionvariants signals
col3Diff = compareLists(list1, list2, 2)

#array of differnces
li = [col1Diff, col2Diff, col3Diff]

#create df out of array and transpose so that it is vertically
newDf = (pd.DataFrame(li)).T
newDf.columns = ['Column1', 'Column2', 'Column3'] #assign columns to the df

df = pd.concat([df3, newDf], axis=0, ignore_index=True) #concetenate df3 with the newDf (df3 is the Inconsistency Sheet)

#write the differences df to the Inconsitency Sheet
with pd.ExcelWriter(filename, mode="a",engine="openpyxl", if_sheet_exists="overlay") as writer:
    df.to_excel(writer, "Inconsistencies", index=False)
