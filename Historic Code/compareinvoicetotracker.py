import csv, math

# compare the fee tracker snapshot to the invoice daybook file
# for this we step through the fee tracker snapshot line by 
# line and look for the corresponding value in the daybook csv
# file.

# setup the variables
job_no_val = str()
fee_type_val = str()
fee_tracker_total_str = str()
ext_fee_val = float()
interco_fee_val = float()
subcon_fee_val = float()
ext_fee_daybook = float()
interco_fee_daybook = float()
subcon_fee_daybook = float()
ext_fee_daybook_sum = int()
interco_fee_daybook_sum = int()
subcon_fee_daybook_sum = int()
ext_fee_daybook_total = int()
interco_fee_daybook_total = int()
subcon_fee_daybook_total = int()
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

# Request the names for the csv files from the user
daybook = input('Enter the file name of the invoice file (invdaybook1.csv)> ')
if daybook == '' : daybook = 'invdaybook1.csv'

snapshot = input('Enter the file name of the snapshot file (feeposition.csv)> ')
if snapshot == '' : snapshot = 'feeposition.csv'
print (snapshot)

# Define the output file
output = 'Invoice vs fee traker.txt'
output_file = open(output, 'w')


# try to open the snapshot_file
try:
    snapshot_file = open(snapshot, 'r')
    csv_snapshot_file = csv.reader(snapshot_file)
    print (snapshot_file, 'used')
except:
    print (snapshot_file, 'does not exist, exiting')
    quit()
    
# try to open the daybook_file
try:
    daybook_file = open(daybook, 'r')
    csv_daybook_file = csv.reader(daybook_file)
    print (daybook_file, 'used')

except:
    print (daybook_file, 'does not exist, exiting')
    quit()

# Step through the fee_snapshot_file a line at a time
# and create a list of the job numbers.
for row in csv_snapshot_file:
    if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
    # read in the first job number
    job_no_val = row[0]
    # check if this job number is in the list
    if job_no_val in job_no_list: continue
    # if the job number is new add it to the list
    job_no_list.append(job_no_val)
unique_jobno_count = len(job_no_list)
jobno_string = ('There are ' + str(unique_jobno_count) + ' job numbers to search\n\n')
print (jobno_string)
output_file.write(jobno_string)

# Now run through the snapshot file again and collated
# all the fee values for the project.
for job_no_val in job_no_list:

    # return to first row of snapshot file and initialise variables
    snapshot_file.seek(0)
    ext_fee_val = 0
    interco_fee_val = 0
    subcon_fee_val = 0

    # now run through the snapshot file looking for job_no
    for row in csv_snapshot_file:
        if not row[0].startswith('qs') or row[0].startswith('pqs') : continue
        # now loop through the fee_snapshot_file and find all corresponding
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
    # the corresponding job in the invoice file and check if they
    # match.
    # first return to the top of the daybook file and initialise variables
    ext_fee_daybook_sum = 0
    interco_fee_daybook_sum = 0
    subcon_fee_daybook_sum = 0
    daybook_file.seek(0)

    # now run through the daybook file looking for job_no
    for row in csv_daybook_file:
        job_start_check = row[3]
        if not job_start_check.startswith('qs') or job_start_check.startswith('pqs') : continue
        # now loop throug the daybook file and return the values
        # for each fee type.
        if job_start_check != job_no_val : continue
        ext_fee_returned = float(row[9])
        interco_fee_returned = float(row[10])
        subcon_fee_returned = float(row[11])
        ext_fee_daybook = math.floor(ext_fee_returned + 0.5)
        interco_fee_daybook = math.floor(interco_fee_returned + 0.5)
        subcon_fee_daybook = math.floor(subcon_fee_returned + 0.5)
        ext_fee_daybook_sum = ext_fee_daybook + ext_fee_daybook_sum
        interco_fee_daybook_sum = interco_fee_daybook +  interco_fee_daybook_sum
        subcon_fee_daybook_sum = subcon_fee_daybook + subcon_fee_daybook_sum
    # collate the values in to job_details_found to compare
    job_details_found = (ext_fee_daybook_sum, interco_fee_daybook_sum,subcon_fee_daybook_sum)
    print ('Found        ',job_no_val, 'with values', job_details_found)
    ext_fee_daybook_total = ext_fee_daybook_sum + ext_fee_daybook_total
    interco_fee_daybook_total = interco_fee_daybook_sum + interco_fee_daybook_total
    subcon_fee_daybook_total = subcon_fee_daybook_sum + subcon_fee_daybook_total
    if job_details_search == job_details_found :
        matched_job = matched_job + 1
        continue
    non_matched_job = non_matched_job + 1
    print ('=' * 80)
    print (job_no_val)
    print ('Fee Tracker', 'External Fee:',ext_fee_val,'Intercompany Fee:',interco_fee_val,'Subcontract Fee', subcon_fee_val)
    print ('Daybook'    , 'External Fee:',ext_fee_daybook_sum,'Intercompany Fee:',interco_fee_daybook_sum,'Subcontract Fee', subcon_fee_daybook_sum)
    print ('=' * 80)
    
    # Write the exceptions to the output_file
    # setup output variables
    ext_fee_nice = format(ext_fee_val, ',')
    interco_fee_nice = format(interco_fee_val, ',')
    subcon_fee_nice = format(subcon_fee_val, ',')
    ext_daybook_nice = format(ext_fee_daybook_sum, ',')
    interco_daybook_nice = format(interco_fee_daybook_sum, ',')
    subcon_daybook_nice = format(subcon_fee_daybook_sum, ',')
    seperate = ('=' * 95)    
    newline = ('\n')
    # write the lines to the file
    output_file.write(seperate)
    output_file.write(newline)
    output_file.write(job_no_val + '  -  ' + job_description + '  -  ' + job_owner)
    output_file.write(newline)
    output_file.write(newline)
    output_file.write('Fee Tracker - External Fee: ' + ext_fee_nice)
    output_file.write(' ' * (8 - len(ext_fee_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_fee_nice)
    output_file.write(' ' * (8 - len(interco_fee_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_fee_nice)
    output_file.write(newline)                     
    output_file.write('Daybook     - External Fee: ' + ext_daybook_nice)
    output_file.write(' ' * (8 - len(ext_daybook_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_daybook_nice)
    output_file.write(' ' * (8 - len(interco_daybook_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_daybook_nice)
    output_file.write(newline)


# Now run through the invoice file and check if there are any job numbers
# that are not in the snapshot file.

# return to the start of the invoice file
daybook_file.seek(0)

for row in csv_daybook_file:
    job_start_check = row[3]
    if not job_start_check.startswith('qs') or job_start_check.startswith('pqs') : continue
    # read in the first job number
    job_no_inv = row[3]
    # check if this job number is in the list
    if job_no_inv in job_no_inv_list: continue
    # if the job number is new add it to the list
    job_no_inv_list.append(job_no_inv)

# Check the daybook job number list against the snapshot job number list and create a list of exceptions
for job_check in job_no_inv_list:
    if job_check in job_no_list : continue
    new_inv_job_list.append(job_check)
new_inv_job_count = len(new_inv_job_list)
if new_inv_job_count > 0:
    new_inv_job_list_string = ('The following ' + str(new_inv_job_count) + ' jobs are in the daybook file but not on the fee tracker\n\n')
    print (new_inv_job_list_string)
    output_file.write(seperate)
    output_file.write(newline * 2)
    output_file.write(new_inv_job_list_string)

# Use the list of exceptions and get the new invoice jobs from the invoice file
daybook_file.seek(0)

for row in csv_daybook_file:
    job_start_check = row[3]
    if not job_start_check.startswith('qs') or job_start_check.startswith('pqs') : continue
    # read in the job number
    job_no_inv = row[3]
    # check if this job is in the list of exceptions
    if not job_no_inv in new_inv_job_list : continue
    job_description = row[4]
    ext_fee_returned = float(row[9])
    interco_fee_returned = float(row[10])
    subcon_fee_returned = float(row[11])
    ext_fee_daybook = math.floor(ext_fee_returned + 0.5)
    interco_fee_daybook = math.floor(interco_fee_returned + 0.5)
    subcon_fee_daybook = math.floor(subcon_fee_returned + 0.5)
    ext_fee_daybook_sum = ext_fee_daybook + ext_fee_daybook_sum
    interco_fee_daybook_sum = interco_fee_daybook +  interco_fee_daybook_sum
    subcon_fee_daybook_sum = subcon_fee_daybook + subcon_fee_daybook_sum
    ext_fee_daybook_total = ext_fee_daybook_sum + ext_fee_daybook_total
    interco_fee_daybook_total = interco_fee_daybook_sum + interco_fee_daybook_total
    subcon_fee_daybook_total = subcon_fee_daybook_sum + subcon_fee_daybook_total
    # collate the values in to job_details_found to compare
    job_details_found = (ext_fee_daybook_sum, interco_fee_daybook_sum,subcon_fee_daybook_sum)
    print ('Found        ',job_no_inv, 'with values', job_details_found)
    print ('=' * 80)
    print (job_no_inv)
    print ('Daybook - Additional Fee'    , 'External Fee:',ext_fee_daybook_sum,'Intercompany Fee:',interco_fee_daybook_sum,'Subcontract Fee', subcon_fee_daybook_sum)
    print ('=' * 80)

    # Write the exceptions to the output_file
    # setup output variables
    ext_daybook_nice = format(ext_fee_daybook_sum, ',')
    interco_daybook_nice = format(interco_fee_daybook_sum, ',')
    subcon_daybook_nice = format(subcon_fee_daybook_sum, ',')
    seperate = ('=' * 95)    
    newline = ('\n')
    # write the lines to the file
    output_file.write(seperate)
    output_file.write(newline)
    output_file.write(job_no_inv + '  -  ' + job_description)
    output_file.write(newline)
    output_file.write(newline)
    output_file.write('Daybook     - External Fee: ' + ext_daybook_nice)
    output_file.write(' ' * (8 - len(ext_daybook_nice)))
    output_file.write('|  Intercompany Fee: ' + interco_daybook_nice)
    output_file.write(' ' * (8 - len(interco_daybook_nice)))
    output_file.write('|  Subcontract Fee: ' + subcon_daybook_nice)
    output_file.write(newline)


# Output the fee tracker closing statements and totals
ext_fee_total_nice = format(ext_fee_total, ',')
interco_fee_total_nice = format(interco_fee_total, ',')
subcon_fee_total_nice = format(subcon_fee_total, ',')
fee_tracker_total = format(ext_fee_total + interco_fee_total + subcon_fee_total, ',')
fee_tracker_charge = format(ext_fee_total + interco_fee_total, ',')
output_file.write(seperate)
output_file.write(newline * 2)
fee_tracker_sum_str = ('External Fee: ' + str(ext_fee_total_nice) + '    Intercompany Fee: ' + str(interco_fee_total_nice) + '    Intercompany Sub-con Fee: ' + str(subcon_fee_total_nice))
fee_tracker_total_str = ('Total: ' + fee_tracker_total)
fee_tracker_charge_str = ('Total Chargeable: ' + fee_tracker_charge)
print('Fee Tracker Totals')
print(fee_tracker_sum_str)
print(fee_tracker_total_str)
print(fee_tracker_charge_str)
print (seperate)
output_file.write('Fee Tracker Totals\n')
output_file.write(fee_tracker_sum_str)
output_file.write(newline)
output_file.write(fee_tracker_total_str)
output_file.write(newline)
output_file.write(fee_tracker_charge_str)
output_file.write(newline * 2)

# Output the daybook closing statements and totals
ext_fee_daybook_total_nice = format(ext_fee_daybook_total, ',')
interco_fee_daybook_total_nice = format(interco_fee_daybook_total, ',')
subcon_fee_daybook_total_nice = format(subcon_fee_daybook_total, ',')
daybook_total = format(ext_fee_daybook_total + interco_fee_daybook_total + subcon_fee_daybook_total, ',')
daybook_charge = format(ext_fee_daybook_total + interco_fee_daybook_total, ',')
output_file.write(seperate)
output_file.write(newline * 2)
inv_daybook_sum_str = ('External Fee: ' + str(ext_fee_daybook_total_nice) + '    Intercompany Fee: ' + str(interco_fee_daybook_total_nice) + '    Intercompany Sub-con Fee: ' + str(subcon_fee_daybook_total_nice))
inv_daybook_total_str = ('Total: ' + daybook_total)
inv_daybook_charge_str = ('Total Chargeable: ' + daybook_charge)
print('Invoice Daybook Totals')
print(inv_daybook_sum_str)
print(inv_daybook_total_str)
print(inv_daybook_charge_str)
output_file.write('Invoice Daybook Totals\n')
output_file.write(inv_daybook_sum_str)
output_file.write(newline)
output_file.write(inv_daybook_total_str)
output_file.write(newline)
output_file.write(inv_daybook_charge_str)
output_file.write(newline * 2)
output_file.write(seperate)
output_file.write(newline)

print('Matched ' + str(matched_job) + ' jobs')
print ('There were ' + str(non_matched_job) + ' that were incorrect on fee tracker')
if matched_job + non_matched_job == unique_jobno_count:
    print('Success, all of the jobs were found in the daybook')
    output_file.write('Success, all of the jobs were found in the daybook\n\n')
elif matched_job + non_matched_job == unique_jobno_count:
    print('CAUTION, there are missing jobs on the fee tracker')
    output_file.write('CAUTION, there are missing jobs on the fee tracker\n')
output_file.write('Matched ' + str(matched_job + non_matched_job) + ' jobs\n\n')
output_file.write('There were ' + str(non_matched_job) + ' that were incorrect on fee tracker\n\n')
if new_inv_job_count > 0:
    output_file.write('The are ' + str(new_inv_job_count) + ' jobs are in the daybook file but not on the fee tracker\n\n')
# Close the output file      
output_file.close()



