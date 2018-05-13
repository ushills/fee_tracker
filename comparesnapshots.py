# import libraries for GUI
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.simpledialog import askstring

import csv
import math
import time

# remove tkinter window
root = tk.Tk()
root.withdraw()

# Select the worksheet that we want to read
snapshotDir = r'C:\Users\hillian\Documents\Dropbox\Fee Tracker Python'
outputFileDir = r'C:\Users\hillian\Documents\Dropbox\Fee Tracker Python'

# Get the current date
date = time.strftime("%d%m%y")
print (date)

# compare the fee tracker snapshot1 to the invoice snapshot2 file
# for this we step through the fee tracker snapshot1 line by 
# line and look for the corresponding value in the snapshot2 csv
# file.

# setup the variables
job_no_val = str()
fee_type_val = str()
ext_fee_val = float()
interco_fee_val = float()
subcon_fee_val = float()
ext_fee2_val = float()
interco_fee2_val = float()
subcon_fee2_val = float()
ext_fee_snapshot2 = float()
interco_fee_snapshot2 = float()
subcon_fee_snapshot2 = float()
ext_fee2_total = float()
interco_fee2_total = float()
subcon_fee2_total = float()
ext_fee_snapshot2_total = float()
interco_fee_snapshot2_total = float()
subcon_fee_snapshot2_total = float()
non_matched_job = int()
matched_job = int()
job_check = int()
new_inv_job_no_count = int()
job_no_list = list()
job_no_inv_list = list()
new_inv_job_list = list()
job_details_search = list()
job_details_found = list ()
ext_fee_total = int()
interco_fee_total = int()
subcon_fee_total = int()
ext_fee2_total = int()
interco_fee2_total = int()
subcon_fee2_total = int()

# Request the names for the csv files from the user
#snapshot1 = input('Enter the file name of the first snapshot1 file (snapshot1.csv)> ')
snapshot1 = filedialog.askopenfilename(filetypes = (("csv files", "*.csv"),("All files", "*.*")), title = ("Select the first snapshot file"), initialdir = snapshotDir)
if snapshot1 == '' : snapshot1 = 'snapshot1.csv'
snapshot1_filename = ((snapshot1.split("/"))[-1])
snapshot1_filename = (snapshot1_filename.split("."))[0]

#snapshot2 = input('Enter the file name of the second snapshot1 file (snapshot2.csv)> ')
snapshot2 = filedialog.askopenfilename(filetypes = (("csv files", "*.csv"),("All files", "*.*")), title = ("Select the second snapshot file"), initialdir = snapshotDir)
if snapshot2 == '' : snapshot2 = 'snapshot2.csv'
snapshot2_filename = ((snapshot2.split("/"))[-1])
snapshot2_filename = (snapshot2_filename.split("."))[0]

# Define the output file
# Request name of the output file name
outputFile = 'snapshot comparison ' + date + '.txt'
string = 'Enter the file name of the output text file(', outputFile,')>'
#csvFileName = input(string)
output = filedialog.asksaveasfilename(title = "Output file name", defaultextension=".txt", initialfile=outputFile, initialdir = outputFileDir)
if output == '':
    output = outputFile
output_file = open(output, 'w')


# try to open the snapshot1_file
try:
    snapshot1_file = open(snapshot1, 'r')
    csv_snapshot1_file = csv.reader(snapshot1_file)
    print (snapshot1_filename, 'used')
except:
    print (snapshot1_filename, 'does not exist, exiting')
    quit()
    
# try to open the snapshot2_file
try:
    snapshot2_file = open(snapshot2, 'r')
    csv_snapshot2_file = csv.reader(snapshot2_file)
    print (snapshot2_filename, 'used')

except:
    print (snapshot2_filename, 'does not exist, exiting')
    quit()

# Step through the fee_snapshot1_file a line at a time
# and create a list of the job numbers.
for row in csv_snapshot1_file:
    if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
    # read in the first job number
    job_no_val = row[0]
    # check if this job number is in the list
    if job_no_val in job_no_list: continue
    # if the job number is new add it to the list
    job_no_list.append(job_no_val)
unique_jobno_count = len(job_no_list)
jobno_string = ('There are ' + str(unique_jobno_count) + ' job numbers to compare\n\n')
print (jobno_string)
output_file.write(jobno_string)

# Now run through the snapshot1 file again and collated
# all the fee values for the project.
for job_no_val in job_no_list:

    # return to first row of snapshot1 file and initialise variables
    snapshot1_file.seek(0)
    ext_fee_val = 0
    interco_fee_val = 0
    subcon_fee_val = 0

    # now run through the snapshot1 file looking for job_no
    for row in csv_snapshot1_file:
        if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
        # now loop through the fee_snapshot1_file and find all corresponding
        # values for the various fee types
        # read the fee type and then set the value depending on the fee type
        if not row[0] == job_no_val : continue
        job_description = row[1]
        job_owner = row[4]
        fee_type_val = row[3]
        fee_val = float(row[5])
        fee_val = math.floor(fee_val + 0.5)
        if fee_type_val == 'External Fee' : ext_fee_val = fee_val + ext_fee_val
        if fee_type_val == 'Inter-Company Fee' : interco_fee_val = fee_val + interco_fee_val
        if fee_type_val == 'Inter-Co Sub-Contract' : subcon_fee_val = fee_val + subcon_fee_val
    ext_fee_total = ext_fee_val + ext_fee_total
    interco_fee_total = interco_fee_val + interco_fee_total
    subcon_fee_total = subcon_fee_val + subcon_fee_total
    job_details_search = (ext_fee_val,interco_fee_val,subcon_fee_val)
    print ('Searching for',job_no_val, 'with values', job_details_search)

    # now for the job number we have just collated we need to find
    # the corresponding job in the snapshot2 file and check if they
    # match.
    # first return to the top of the snapshot2 file and initialise variables
    snapshot2_file.seek(0)
    ext_fee2_val = 0
    interco_fee2_val = 0
    subcon_fee2_val = 0

    # now run through the snapshot2 file looking for job_no
    for row in csv_snapshot2_file:
        if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
        # now loop through the snapshot2 file and return the values
        # for each fee type.
        if row[0] != job_no_val : continue
        fee2_type_val = row[3]
        fee2_val = float(row[5])
        fee2_val = math.floor(fee2_val + 0.5)
        if fee2_type_val == 'External Fee' : ext_fee2_val = fee2_val + ext_fee2_val
        if fee2_type_val == 'Inter-Company Fee' : interco_fee2_val = fee2_val + interco_fee2_val
        if fee2_type_val == 'Inter-Co Sub-Contract' : subcon_fee2_val = fee2_val + subcon_fee2_val
    ext_fee2_total = ext_fee2_val + ext_fee2_total
    interco_fee2_total = interco_fee2_val +  interco_fee2_total
    subcon_fee2_total = subcon_fee2_val + subcon_fee2_total
    # collate the values in to job_details_found to compare
    job_details_found = (ext_fee2_val, interco_fee2_val,subcon_fee2_val)
    print ('Found        ',job_no_val, 'with values', job_details_found)
    if job_details_search == job_details_found :
        matched_job = matched_job + 1
        continue
    non_matched_job = non_matched_job + 1
    print ('=' * 80)
    print (job_no_val)
    print (snapshot1_filename,' External Fee:',ext_fee_val,'Intercompany Fee:',interco_fee_val,'Subcontract Fee', subcon_fee_val)
    print (snapshot2_filename,' External Fee:',ext_fee2_val,'Intercompany Fee:',interco_fee2_val,'Subcontract Fee', subcon_fee2_val)
    print ('=' * 80)
    
    # Write the exceptions to the output_file
    # setup output variables
    ext_fee_nice = format(ext_fee_val, ',')
    interco_fee_nice = format(interco_fee_val, ',')
    subcon_fee_nice = format(subcon_fee_val, ',')
    ext_fee2_nice = format(ext_fee2_val, ',')
    interco_fee2_nice = format(interco_fee2_val, ',')
    subcon_fee2_nice = format(subcon_fee2_val, ',')
    seperate = ('=' * 95)    
    newline = ('\n')
    # write the lines to the file
    output_file.write(seperate)
    output_file.write(newline)
    output_file.write(job_no_val + '  -  ' + job_description + '  -  ' + job_owner)
    output_file.write(newline)
    output_file.write(newline)
    output_file.write(snapshot1_filename + ' - External Fee: ' + ext_fee_nice)
    output_file.write(' ' * (8 - len(ext_fee_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_fee_nice)
    output_file.write(' ' * (8 - len(interco_fee_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_fee_nice)
    output_file.write(newline)                     
    output_file.write(snapshot2_filename + ' - External Fee: ' + ext_fee2_nice)
    output_file.write(' ' * (8 - len(ext_fee2_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_fee2_nice)
    output_file.write(' ' * (8 - len(interco_fee2_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_fee2_nice)
    output_file.write(newline)


# Now run through the snapshot file and check if there are any job numbers
# that are not in the snapshot1 file.

# return to the start of the invoice file
snapshot2_file.seek(0)

for row in csv_snapshot2_file:
    if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
    # read in the first job number
    job_no_inv = row[0]
    # check if this job number is in the list
    if job_no_inv in job_no_inv_list: continue
    # if the job number is new add it to the list
    job_no_inv_list.append(job_no_inv)

# Check the snapshot2 job number list against the snapshot1 job number list and create a list of exceptions
for job_check in job_no_inv_list:
    if job_check in job_no_list : continue
    new_inv_job_list.append(job_check)
new_inv_job_count = len(new_inv_job_list)
new_inv_job_list_string = ('The following ' + str(new_inv_job_count) + ' jobs are in the ' + snapshot2_filename + ' file but not in ' + snapshot1_filename + '\n\n')
print (new_inv_job_list_string)
print (new_inv_job_list)
output_file.write(seperate)
output_file.write(newline * 2)
output_file.write(new_inv_job_list_string)

# Use the list of exceptions and get the new invoice jobs from the invoice file


for new_job in new_inv_job_list:
    snapshot2_file.seek(0)
    ext_fee2_val = 0
    interco_fee2_val = 0
    subcon_fee2_val = 0
    for row in csv_snapshot2_file:
        if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
        # check if this job is in the list of exceptions
        if row[0] != new_job : continue
        print (row[0])
        job_description = row[1]
        fee2_type_val = row[3]
        fee2_val = float(row[5])
        fee2_val = math.floor(fee2_val + 0.5)
        if fee2_type_val == 'External Fee' : ext_fee2_val = fee2_val + ext_fee2_val
        if fee2_type_val == 'Inter-Company Fee' : interco_fee2_val = fee2_val + interco_fee2_val
        if fee2_type_val == 'Inter-Co Sub-Contract' : subcon_fee2_val = fee2_val + subcon_fee2_val
    ext_fee2_total = ext_fee2_val + ext_fee2_total
    interco_fee2_total = interco_fee2_val +  interco_fee2_total
    subcon_fee2_total = subcon_fee2_val + subcon_fee2_total
    # collate the values in to job_details_found to compare
    job_details_found = (ext_fee2_val, interco_fee2_val,subcon_fee2_val)
    print ('Found        ',job_no_inv, 'with values', job_details_found)
    print ('=' * 80)
    print (new_job)
    print (snapshot2_filename, 'External Fee:',ext_fee2_val,'Intercompany Fee:',interco_fee2_val,'Subcontract Fee', subcon_fee2_val)
    print ('=' * 80)

    # Write the exceptions to the output_file
    # setup output variables
    ext_fee2_nice = format(ext_fee2_val, ',')
    interco_fee2_nice = format(interco_fee2_val, ',')
    subcon_fee2_nice = format(subcon_fee2_val, ',')
    seperate = ('=' * 95)    
    newline = ('\n')
    # write the lines to the file
    output_file.write(seperate)
    output_file.write(newline)
    output_file.write(new_job + '  -  ' + job_description)
    output_file.write(newline)
    output_file.write(newline)
    output_file.write(snapshot2_filename + ' - External Fee: ' + ext_fee2_nice)
    output_file.write(' ' * (8 - len(ext_fee2_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_fee2_nice)
    output_file.write(' ' * (8 - len(interco_fee2_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_fee2_nice)
    output_file.write(newline)


# Output the snapshot1 totals
ext_fee_total_nice = format(ext_fee_total, ',')
interco_fee_total_nice = format(interco_fee_total, ',')
subcon_fee_total_nice = format(subcon_fee_total, ',')
snapshot1_total = format(ext_fee_total + interco_fee_total + subcon_fee_total, ',')
snapshot1_charge = format(ext_fee_total + interco_fee_total, ',')
output_file.write(seperate)
output_file.write(newline * 2)
snapshot1_sum_str = ('External Fee: ' + str(ext_fee_total_nice) + '    Intercompany Fee: ' + str(interco_fee_total_nice) + '    Intercompany Sub-con Fee: ' + str(subcon_fee_total_nice))
snapshot1_total_str = ('Total: ' + snapshot1_total)
snapshot1_charge_str = ('Total chargeable: ' + snapshot1_charge)
print(snapshot1_filename + ' totals')
print(snapshot1_sum_str)
print(snapshot1_total_str)
print(snapshot1_charge_str)
print (seperate)
output_file.write(snapshot1_filename + ' totals\n')
output_file.write(snapshot1_sum_str)
output_file.write(newline)
output_file.write(snapshot1_total_str)
output_file.write(newline)
output_file.write(snapshot1_charge_str)
output_file.write(newline * 2)

# Output the snapshot2 totals
ext_fee2_total_nice = format(ext_fee2_total, ',')
interco_fee2_total_nice = format(interco_fee2_total, ',')
subcon_fee2_total_nice = format(subcon_fee2_total, ',')
snapshot2_total = format(ext_fee2_total + interco_fee2_total + subcon_fee2_total, ',')
snapshot2_charge = format(ext_fee2_total + interco_fee2_total, ',')
output_file.write(seperate)
output_file.write(newline * 2)
snapshot2_sum_str = ('External Fee: ' + str(ext_fee2_total_nice) + '    Intercompany Fee: ' + str(interco_fee2_total_nice) + '    Intercompany Sub-con Fee: ' + str(subcon_fee2_total_nice))
snapshot2_total_str = ('Total: ' + snapshot2_total)
snapshot2_charge_str = ('Total chargeable: ' + snapshot2_charge)
print(snapshot2_filename + ' totals')
print(snapshot2_sum_str)
print(snapshot2_total_str)
print(snapshot2_charge_str)
print(seperate)
output_file.write(snapshot2_filename +' totals\n')
output_file.write(snapshot2_sum_str)
output_file.write(newline)
output_file.write(snapshot2_total_str)
output_file.write(newline)
output_file.write(snapshot2_charge_str)
output_file.write(newline * 2)
output_file.write(seperate)

# Output the deltas 
ext_delta_total_nice = format(ext_fee2_total - ext_fee_total, ',')
interco_delta_total_nice = format(interco_fee2_total - interco_fee_total, ',')
subcon_delta_total_nice = format(subcon_fee2_total - subcon_fee_total, ',')
delta_total = format((ext_fee2_total + interco_fee2_total + subcon_fee2_total) - (ext_fee_total + interco_fee_total + subcon_fee_total), ',')
delta_charge = format((ext_fee2_total + interco_fee2_total) - (ext_fee_total + interco_fee_total), ',')
output_file.write(newline * 2)
delta_sum_str = ('External Fee: ' + str(ext_delta_total_nice) + '    Intercompany Fee: ' + str(interco_delta_total_nice) + '    Intercompany Sub-con Fee: ' + str(subcon_delta_total_nice))
delta_total_str = ('Total: ' + delta_total)
delta_charge_str = ('Total chargeable: ' + delta_charge)
print('Delta ',snapshot1_filename,' vs ',snapshot2_filename,' totals')
print(delta_sum_str)
print(delta_total_str)
print(delta_charge_str)
print(seperate)
output_file.write('Delta ' + snapshot1_filename + ' vs ' + snapshot2_filename + ' totals\n')
output_file.write(delta_sum_str)
output_file.write(newline)
output_file.write(delta_total_str)
output_file.write(newline)
output_file.write(delta_charge_str)
output_file.write(newline * 2)
output_file.write(seperate)
output_file.write(newline * 2)

# Output the closing statements
print('Matched ' + str(matched_job) + ' jobs')
print ('There were ' + str(non_matched_job) + ' that were incorrect on fee tracker')
if matched_job + non_matched_job == unique_jobno_count:
    print('Success, all of the jobs in ', snapshot1_filename, ' were found in ', snapshot2_filename)
    output_file.write('Success, all of the jobs in ' + snapshot1_filename + ' were found in ' + snapshot2_filename)
    output_file.write(newline * 2)
elif matched_job + non_matched_job == unique_jobno_count:
    print('CAUTION, there are missing jobs on the fee tracker')
    output_file.write('CAUTION, there are missing jobs on the in ' + snapshot2_filename + '\n')
    output_file.write(newline * 2)
output_file.write('Matched ' + str(matched_job + non_matched_job) + ' jobs\n\n')
output_file.write('There are ' + str(non_matched_job) + ' that have changed between snapshots\n\n')
output_file.write('The are ' + str(new_inv_job_count) + ' new jobs are in the that were not in the previous snapshot\n\n')
# Close the output file      
output_file.close()



