date /t
@echo off
call activate
echo %1
python C:\Users\Engineer3\Documents\timesheet_manager\payroll_fn.py %1
call conda deactivate