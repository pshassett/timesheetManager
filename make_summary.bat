date /t
@echo off
call activate
echo %1
python C:\Users\Engineer3\Documents\timesheet_manager\jobtasksummary_fn.py %1
call deactivate