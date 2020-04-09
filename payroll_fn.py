# -*- coding: utf-8 -*-
"""
Created on Thu Mar 12 13:32:47 2020

@author: Engineer3


Parse the timesheet file and produce payroll information for each employee.


"""

from os import path
import sys
import pandas as pd


# Define the output folder
payroll_folder = path.abspath(r"C:\Users\Engineer3\Desktop\Payroll")

def make_payroll_report(timesheet):
    # Load the data of the spreadsheet as a DataFrame object.
    timesheets = pd.read_excel(timesheet, sheet_name=None)
    # Create our output dataframe templates
    period1 = pd.DataFrame(columns=['Start Date', 'End Date', 'Overtime', 'Regular', 'Vacation', 'Sick', 'Holiday', 'Total'])
    period2 = pd.DataFrame(columns=['Start Date', 'End Date', 'Overtime', 'Regular', 'Vacation', 'Sick', 'Holiday', 'Total'])
    for employee in list(timesheets.keys())[1:]:
        print(employee)
        # Get just the data for this employee.
        empdata = timesheets[employee]
        # Fix the dates to be just dates, no timestamp.
        empdata.Date
        # Remove the top template row from the data
        empdata.drop([0], inplace=True)
        # Remove rows that are missing hours or job number.
        empdata.dropna(axis='rows',subset=['Job No.', 'Hours'], inplace=True)        
        # Determine the index of the last entry in the first pay period.
        index = 0
        for date in empdata.Date:
            if date.day >= 16:
                break
            else:
                index +=1
        # Set the date bounds for the two time periods.
        bounds = [[0, index], [index, -1]]
        # Create payroll summary reports for both pay periods.        
        for bound in bounds:
            summary = {}
            # Select the data just for this pay period.
            selection = empdata.iloc[bound[0]:bound[1],:]
            # Get the start and end dates of the pay period.
            summary['Start Date'] = selection.Date.iloc[0]
            summary['End Date'] = selection.Date.iloc[-1]
            # Get the hour sums for the necessary categories.
            summary['Overtime'] = selection.Overtime.sum()
            summary['Total'] = selection.Hours.sum()
            summary['Vacation'] = selection.loc[selection["Job No."]  == "VAC"].Hours.sum()
            summary['Sick'] = selection.loc[selection["Job No."]  == "SICK"].Hours.sum()
            summary['Holiday'] = selection.loc[selection["Job No."]  == "HOL"].Hours.sum()
            summary['Regular'] = summary['Total'] - (summary['Vacation'] + summary['Sick'] + summary['Holiday'] + summary['Overtime'])
            # Convert the summary dict to a pd.Series and append to the appropriate payperiod df.
            summary = pd.Series(summary)
            summary.name = employee
            if bound[0] == 0:
                period1 = period1.append(summary)
            else:
                period2 = period2.append(summary)
    # Determine the month and year of the data to define the output dir.
    month = summary['End Date'].month_name()
    year = summary['End Date'].year
    month_year = str(month) + str(year)
    output_file = path.join(payroll_folder, month_year + "Payroll.xlsx")
    # Sort both summaries by name.
    period1.sort_index(inplace=True)
    period2.sort_index(inplace=True)
    # Dump both pay period df's to a single .xlsx file.
    with pd.ExcelWriter(output_file) as writer:
        period1.to_excel(writer, sheet_name='Pay Period 1')
        period2.to_excel(writer, sheet_name='Pay Period 2')


# Assign the input file argument to a variable.
file = path.abspath(str(sys.argv[1]))
# Call the summary writer function
make_payroll_report(file)
