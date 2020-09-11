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
        # Remove the template row if it is still there.
        if str(empdata.iloc[0].Date) == '2000-01-01 00:00:00':
            empdata = empdata.iloc[1:]
        # Remove any rows full of NaNs
        empdata.dropna(axis='index', how='all', inplace=True)
        # Skip any timesheets that are empty.
        if len(list(empdata.values)) == 0:
            continue
        # Remove rows that are missing hours or job number.
        # empdata.dropna(axis='rows',subset=['Job No.', 'Hours'], inplace=True)        
        # Reset the indices for sorting. 
        empdata.reset_index(drop=True, inplace=True)
        # Determine the index of the last entry in the first pay period.
        index = 0
        last_index = empdata.Date.last_valid_index()
        if empdata.Date.iloc[last_index].day <= 15:
            index = last_index + 1
        elif empdata.Date.iloc[last_index].day > 15:
            for date in empdata.Date:
                if date.day >= 16:
                    break
                else:
                    index += 1
        # Set the date bounds for the two time periods.
        bounds = [[0, index], [index, None]]
        # Create payroll summary reports for both pay periods.        
        for bound in bounds:
            if bound[0] > last_index:
                summary = {}
                summary['Start Date'] = 0
                summary['End Date'] = 0
                summary['Overtime'] = 0
                summary['Total'] = 0
                summary['Vacation'] = 0
                summary['Sick'] = 0
                summary['Holiday'] = 0
                summary['Regular'] = 0
            else:
                summary = {}
                # Select the data just for this pay period.
                ### Put a try except block here?
                selection = empdata.iloc[bound[0]:bound[1]]
                # Get the start and end dates of the pay period.
                summary['Start Date'] = selection.Date.iloc[0].date().isoformat()
                summary['End Date'] = selection.Date.iloc[-1].date().isoformat()
                # Determine the month and year of the data to define the output dir.
                month = selection.Date.iloc[0].month_name()
                year = selection.Date.iloc[0].year
                month_year = str(month) + str(year)
                # Get the hour sums for the necessary categories.
                summary['Overtime'] = selection.Overtime.sum()
                summary['Total'] = selection.Hours.sum()
                # Isolate the Job no. column and make all str entires upper
                try:
                    summary['Vacation'] = selection.loc[selection["Job No."].str.upper()  == "VAC"].Hours.sum()
                    summary['Sick'] = selection.loc[selection["Job No."].str.upper()  == "SICK"].Hours.sum()
                    summary['Holiday'] = selection.loc[selection["Job No."].str.upper()  == "HOL"].Hours.sum()
                except AttributeError as e:  # If there are no str entries
                    summary['Vacation'] = 0
                    summary['Sick'] = 0
                    summary['Holiday'] = 0
                summary['Regular'] = summary['Total'] - (summary['Vacation'] + summary['Sick'] + summary['Holiday'] + summary['Overtime'])
            # Convert the summary dict to a pd.Series and append to the appropriate payperiod df.
            summary = pd.Series(summary)
            summary.name = employee
            if bound[0] == 0:
                period1 = period1.append(summary)
            else:
                period2 = period2.append(summary)
    output_file = path.join(payroll_folder, month_year + "Payroll.xlsx")
    # Sort both summaries by name.
    period1.sort_index(inplace=True)
    period2.sort_index(inplace=True)
    # Name each pay period.
    period1.name = 'Period1'
    period2.name = 'Period2'
    return period1, period2, output_file
    

def write_payroll_to_file(file, *pay_periods):
    # Dump any amount pay period df's to a single .xlsx file.
    with pd.ExcelWriter(file) as writer:
        for period in pay_periods:
            period.to_excel(writer, sheet_name=period.name)

if __name__ == "__main__":
    # Assign the input file argument to a variable.
    try:
        timesheet = path.abspath(str(sys.argv[1]))
         # Call the summary writer function
        period1, period2, file = make_payroll_report(timesheet)
        write_payroll_to_file(file, period1, period2)
    except IndexError as e:
        print(e)
        #    --------FOR DEBUGGING--------
        timesheet = r'C:\Users\Engineer3\Desktop\HE_Files\Timesheets\GROUPTIMSH2008.xlsm'
        #    -----------------------------
        period1, period2, file = make_payroll_report(timesheet)
        #write_payroll_to_file(file, period1, period2)
    
