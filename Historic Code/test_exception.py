checked = str()

def check_exceptions(to_check):
    if to_check == 'qs????1':
        return 'qs99998'
    elif to_check == 'qs????2':
        return 'qs99998'
    elif to_check == 'qs????3':
        return 'qs99998'
    return to_check

job_no_val = 'qs????1'
job_no_val = check_exceptions(job_no_val)
print (job_no_val)
