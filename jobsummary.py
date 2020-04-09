# -*- coding: utf-8 -*-
"""
Created on Mon Dec 23 16:08:41 2019

@author: Engineer3
"""
from os import path
import pandas as pd

# Open the most recent time sheet file and load the data with pandas.
file = path.abspath(r"C:\Users\Engineer3\Desktop\HE_Files\Timesheets\GROUPTIMSH1912copy.xlsm")
output_folder = path.abspath(r"C:\Users\Engineer3\Desktop\HE Files\JobSummaries")
timesheets = pd.read_excel(file, sheet_name=None)

# Parse the jobs that each employee worked on and enter their hours and tasks into that job's summary
jobs = {}
for employee in list(timesheets.keys())[1:]:
    print(employee)
    # Get just the data for this employee.
    empdata = timesheets[employee]
    # Remove the top template row from the data
    empdata.drop([0], inplace=True)
    # Remove rows that are missing hours or job number.
    empdata.dropna(axis='rows',subset=['Job No.', 'Hours'], inplace=True)
    # Add a column for the employee's name.
    empdata['Employee'] = employee
    # Select the entries for each job and append to that job's summary.
    for job in set(empdata.Job.values):
        if job not in jobs.keys():
            jobs[job] = pd.DataFrame()
            jobs[job] = jobs[job].append(empdata.loc[empdata.Job==job])
        else:
            jobs[job] = jobs[job].append(empdata.loc[empdata.Job==job])
'''
# Sort the data entries for each job by date, then dump the summary to a .xlsx.          
for job, data in jobs.items():
    data.sort_values('Date', inplace=True)
    data.to_excel(path.join(output_folder, str(job)+".xlsx"), index=False)
'''
    