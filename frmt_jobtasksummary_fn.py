# -*- coding: utf-8 -*-
"""
Created on Mon Dec 23 16:08:41 2019

@author: Engineer3
"""
from os import path, makedirs
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Define the summary folder.
summary_folder = path.abspath(r"C:\Users\Engineer3\Desktop\HE_Files\Timesheets\JobSummaries")
def make_job_summary_report(timesheet):
    timesheets = pd.read_excel(timesheet, sheet_name=None)
    # Parse the jobs that each employee worked on and enter their hours and tasks into that job's summary.
    jobs = {}
    for employee in list(timesheets.keys())[1:]:
        print(employee)
        # Get just the data for this employee.
        empdata = timesheets[employee]
        # Remove the top template row from the data.
        empdata.drop([0], inplace=True)
        # Remove vacation, holiday, and sick leave data.
        empdata = empdata.loc[empdata.Job!='Vacation Time']
        empdata = empdata.loc[empdata.Job!='Holiday']
        empdata = empdata.loc[empdata.Job!='Sick Leave']
        # Remove rows that are missing hours or job number.
        empdata.dropna(axis='rows',subset=['Job No.', 'Hours'], inplace=True)
        # Add a column for the employee's name.
        empdata['Employee'] = employee
        # Fetch a date from the entries.
        if str(empdata.Date.max()) != 'NaT':
            lastday = empdata.Date.max()
        # Select the entries for each job and append to that job's summary.
        for job in set(empdata.Job.values):
            # Add the job to the set of jobs if it doesn't already exist.
            if job not in jobs.keys():
                # Populate with an empty dict of tasks
                jobs[job] = {}
            # Update the each job's task summaries with this employee's data.
            # Get the 'tasks' dict for this job.
            tasks = jobs[job]
            # Isolate the employee's data for just this job.
            job_data = empdata.loc[empdata.Job==job]
            for task in set(job_data.Task.values):
                # Add the task to this job's tasks if it doesn't already exist.
                if task not in tasks.keys():
                    tasks[task] = {}
                # Sum the amount of hours spent on the task.
                task_sum = job_data.loc[job_data.Task==task].Hours.sum()
                tasks[task][employee] = task_sum
            # Replace the job's task summaries with the updated task summaries
            jobs[job] = tasks
    # Grab the year/month from the last employee's data
    month = str(lastday.month_name())
    if lastday.month < 10 :
        monthnum = '0' + str(lastday.month)
    else:
        monthnum = str(lastday.month)
    year = str(lastday.year)
    month_year = month + year
    output_dir = path.join(summary_folder, month_year)
    # Make the output directory if it doesn't exist already
    if not path.exists(output_dir):
        makedirs(output_dir)
    # Report out each job's summary.
    # fancy way
    for jobname in jobs.keys():
        # Load the template workbook
        wb = load_workbook(r'C:\Users\Engineer3\Desktop\HE_Files\Timesheets\JobSummaries\template.xlsx')
        ws = wb.active
        # Convert the dict to a df and fill nans with 0.
        jobdf = pd.DataFrame(jobs[jobname]).T.fillna(0)
        # Dump into the template workbook and save with apropriate naming 
        for r in dataframe_to_rows(jobdf, header=True, index=True):
            ws.write(r)        
        # define the output file.
        output_file = path.join(output_dir, str(jobname) + ' '+ year[2:] + monthnum + '_TEST.xlsx')
        wb.save(output_file)
        wb.close()          
    

if __name__ == "__main__":
    # Assign the input file argument to a variable.
    try:
        # Assign the input file argument to a variable.
        file = path.abspath(str(sys.argv[1]))
        # Call the summary writer function
        make_job_summary_report(file)
    except IndexError as e:
        print(e)
        #    --------FOR DEBUGGING--------
        file = r'C:\Users\Engineer3\Desktop\HE_Files\Timesheets\GROUPTIMSH2007.xlsm'
        #    -----------------------------
        make_job_summary_report(file)        
    


        