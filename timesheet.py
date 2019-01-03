import shutil
import os
import openpyxl as pyxl
import time
import pandas as pd

#initialize timer
project_start_time = time.time()
#create timesheet if non-existant
if not os.path.isfile(timesheetname):
    wb=pyxl.Workbook()
    wb.save(Name of timesheet file)
#check for WBS data
with open txt file:
    if wbs in txt:
        pass
    else:
        add_wbs


def new_timesheet():
    #log elapsed project time to timesheet
    log_elapsed_project_time()
    #send current timesheet to archive if exists
    try:
        os.mkdir(Archive folder)
    except OSError as e:
        print('Archive exists')
        pass
    archive_name = timesheet + current time
    os.rename(timesheet, archive_name)
    shutil.move(archive name, archive folder)
    #create new timesheet excel file
    wb=pyxl.Workbook()
    wb.save(Name of timesheet file)
    #start timer
    reset_project_timer()
    
def reset_project_timer():
    #sets globabl variable project_start_time to current time
    global project_start_time
    project_start_time = time.time()

def get_elapsed_project_time():
    #returns elapsed time on project
    elapsed_time = time.time()- project_start_time
    return elapsed_time

def log_elapsed_project_time():
    #read WBS of current InventorWork
    with open txt file:
        network_activity = readlines(second line)
    #check if today's date in timesheet
    wb = pyxl.load_workbook( timesheet path)
    sheet = wb.get_sheet_by_name('Sheet1')
    
    
    
    
