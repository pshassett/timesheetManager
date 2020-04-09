# -*- coding: utf-8 -*-
"""
Created on Mon Dec 23 16:08:41 2019

@author: Engineer3

A function to parse through employee timesheets and produce a summary for each
job worked on and the tasks associated with that job.

Designed to be called from make_summary.bat with the timesheet file name and 
path as a command line argument eg:
    >>>C:\\path\to\make_summary.bat C:\\path\to\some\timesheet.xlsm
This script can be invoked manually with the command above. Alternatively,
and more conveniently, the time sheets can launch make_summary.bat with a 
button press on dashbord as long as they are macro-enabled (.xlsm).
    
"""
from os import path, makedirs
import sys
import pandas as pd


# Define the output folder
summary_folder = path.abspath(r"C:\Users\Engineer3\Desktop\HE_Files\JobSummaries")

def summary_maker(excel_file):
    # Load the data of the spreadsheet as a DataFrame object.
    timesheets = pd.read_excel(excel_file, sheet_name=None)
    # Parse the jobs that each employee worked on and enter their hours and tasks into that job's summary.
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
    # Sort the data entries for each job by date, then dump the summary to a .xlsx.          
    for job, data in jobs.items():
        data.sort_values('Date', inplace=True)
        # Determine the month and year of the data to define the output dir.
        latest_date = data.Date.max()
        month = latest_date.month_name()
        year = latest_date.year
        month_year = str(month) + str(year)
        output_dir = path.join(summary_folder, month_year)
        output_file = path.join(output_dir, str(job)+".xlsx")
        # Make the output directory if it doesn't exist already
        if not path.exists(output_dir):
            makedirs(output_dir)
        # Write out the job data output_file.
        data.to_excel(output_file, index=False)


# Assign the input file argument to a variable.
file = path.abspath(str(sys.argv[1]))
# Call the summary writer function
summary_maker(file)
