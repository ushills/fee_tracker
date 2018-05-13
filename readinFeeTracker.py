# import libraries for GUI
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.simpledialog import askstring

# import libraries for excel, csv and date
import csv
import xlrd
import time

# remove tkinter window
root = tk.Tk()
root.withdraw()

# variables used
# contstants

# Select the worksheet that we want to read
feeTrackerWorksheetName = 'Tracker'
feeTrackerDir = r'F:\Bir\CM\CM Fee Tracker\14-15'
outputFileDir = r'C:\Users\hillian\Documents\Dropbox\Fee Tracker Python'

# Get the current date
date = time.strftime("%d%m%y")
print (date)

# fee tracker file columns

feeTrackerJobNoCol = 0
feeTrackerDescriptionCol = 1
feeTrackerWorkstageCol = 3
feeTrackerFeeTypeCol = 5
feeTrackerCommissionManagerCol = 9
feeTrackerFeeValueArray = [20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
feeTrackerFeeTotal = 0 # initialise totals
# we will set the fee val column based on the period selected


# Request name of the fee tracker excel file
#feeTracker = input('Enter the file name of the fee tracker file (CM Birmingham Tracker 2014-15.xlsm)> ')
#messagebox.showinfo(title = "Read in Fee Tracker", message = "Select the fee tracker file")
feeTracker = filedialog.askopenfilename(filetypes = (("Excel files", "*.xlsm"),("All files", "*.*")), title = ("Select the fee tracker file"), initialdir = feeTrackerDir)
if feeTracker == '':
    feeTracker = r'F:\Bir\CM\CM Fee Tracker\14-15\CM Birmingham Tracker 2014-15.xlsm'
print (feeTracker, ' selected')

# Try to open the fee tracker file
try:
    feeTrackerWorkbook = xlrd.open_workbook(feeTracker)
except:
    print (feeTracker, ' does not exist, exiting')
    quit()

# Request the Period Number to extract
#period = input('Enter the period that you want to extract> ')
period = askstring("Period", "Enter the period that you want to extract")
if period == '':
    quit()
feeTrackerFeeValueCol = feeTrackerFeeValueArray[int(period) - 1]
print ('Column ', feeTrackerFeeValueCol, 'to be used')
print (period, ' selected')

# Request name of the output file name
defaultcsvFileName = 'P' + period + ' ' + date + '.csv'
string = 'Enter the file name of the output csv file(', defaultcsvFileName,')>'
#csvFileName = input(string)
csvFileName = filedialog.asksaveasfilename(title = "Output file name", defaultextension=".csv", initialfile=defaultcsvFileName, initialdir = outputFileDir)
if csvFileName == '':
    csvFileName = defaultcsvFileName
# open the output file and set as csv
csvOutputFile = open(csvFileName, 'w')
csvOutputFileWriter = csv.writer(csvOutputFile, dialect='excel', delimiter=',', quotechar = '"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
print (csvFileName, ' selected')
# write the headers
outputData = ['Job Number', 'Description', 'Workstage', 'Fee Type', 'Commission Manager', 'Fee Sum']
csvOutputFileWriter.writerow(outputData)
 



# iterate through each row of the fee tracker and print the row
feeTrackerWorksheet = feeTrackerWorkbook.sheet_by_name(feeTrackerWorksheetName)
feeTrackerNumRows = feeTrackerWorksheet.nrows - 1
feeTrackerCurrentRow = -1
while feeTrackerCurrentRow < feeTrackerNumRows:
    feeTrackerCurrentRow += 1
    feeTrackerRow = feeTrackerWorksheet.row(feeTrackerCurrentRow)
    # check that the job number cell contains text
    if not feeTrackerWorksheet.cell_type(feeTrackerCurrentRow, feeTrackerJobNoCol) == 1 : continue
    # now check that the job number in this row begins with 'qs' or 'pqs'
    if not feeTrackerWorksheet.cell_value(feeTrackerCurrentRow,feeTrackerJobNoCol).startswith('qs') or feeTrackerWorksheet.cell_value(feeTrackerCurrentRow,feeTrackerJobNoCol).startswith('pqs'): continue
    #print (feeTrackerWorksheet.cell_type(feeTrackerCurrentRow,feeTrackerJobNoCol))
    # check that the value is not zero
    if feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerFeeValueCol) == 0: continue
    # check that the value is not blank
    if feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerFeeValueCol) == '': continue
    
    # now extract the data from the required cells
    feeTrackerJobNumber = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerJobNoCol)
    feeTrackerDescription = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerDescriptionCol)
    feeTrackerWorkstage = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerWorkstageCol)
    feeTrackerFeeType = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerFeeTypeCol)
    feeTrackerCommissionManager = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerCommissionManagerCol)
    feeTrackerFeeValue = feeTrackerWorksheet.cell_value(feeTrackerCurrentRow, feeTrackerFeeValueCol)
    feeTrackerFeeTotal += feeTrackerFeeValue

    # at this stage we want to write the variables above to the csv file
    outputData = [feeTrackerJobNumber, feeTrackerDescription, feeTrackerWorkstage, feeTrackerFeeType, feeTrackerCommissionManager, feeTrackerFeeValue]
    csvOutputFileWriter.writerow(outputData)
 
    print ('Writing ', feeTrackerJobNumber, feeTrackerDescription, feeTrackerWorkstage, feeTrackerFeeType, feeTrackerCommissionManager, feeTrackerFeeValue)
print (feeTrackerFeeTotal)

# Close the output file
csvOutputFile.close()
    
    


