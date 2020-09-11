# -*- coding: utf-8 -*-
"""
Created on Mon Dec 23 16:08:41 2019

@author: Engineer3
"""
from os import path, makedirs
import pandas as pd

# Open the most recent time sheet file and load the data with pandas.
file = path.abspath(r"C:\Users\Engineer3\Desktop\HE_Files\Timesheets\GROUPTIMSH2004.xlsm")
summary_folder = path.abspath(r"C:\Users\Engineer3\Desktop\HE_Files\JobSummaries")

timesheets = pd.read_excel(file, sheet_name=None)
# Parse the jobs that each employee worked on and enter their hours and tasks into that job's summary
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
    # fetch date
    if str(empdata.Date.max()) != 'NaT':
            lastday = empdata.Date.max()
    # Add a column for the employee's name.
    empdata['Employee'] = employee
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
# lastday = empdata.Date.max()
month = str(lastday.month_name())
monthnum = str(lastday.month)
year = str(lastday.year)
month_year = month + year
output_dir = path.join(summary_folder, month_year)
# Make the output directory if it doesn't exist already
if not path.exists(output_dir):
    makedirs(output_dir)
'''
# Report out each job's summary.
for jobname in jobs.keys():
    # define the output file.
    output_file = path.join(output_dir, str(jobname) + ' '+ year[2:] + monthnum + '.xlsx')
    # Convert the dict to a df and fill nans with 0.
    jobdf = pd.DataFrame(jobs[jobname]).T.fillna(0)
    jobdf.to_excel(output_file)
'''