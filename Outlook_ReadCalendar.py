# Import from packages
import os
import argparse
import logging
from datetime import datetime, timedelta, time
import win32com.client

# Import from modules
from MyCls import initialise_app, finalise_app, handle_exception, parse_datetime

# Initialise project
CURR_DIR, CURR_FILE = os.path.split(__file__)
PROJ_NAME = CURR_FILE.split('.')[0]

# Get command line arguments
my_arg_parser = argparse.ArgumentParser(description=f"{PROJ_NAME}")
#my_arg_parser.add_argument("arg1", help="Text1")
my_arg_parser.add_argument("--log", help="DEBUG to enter debug mode")
args = my_arg_parser.parse_args()

# Initialise app
initialise_app(PROJ_NAME, args.log)
logger = logging.getLogger("my_logger")

# # Get environment variables
# env_var1 = os.getenv("env_var1")
# env_var2 = os.getenv("env_var2")
# if env_var1 == None or env_var2 == None:
#     handle_exception("Missing environment variables!")

##################################################
# Variables
##################################################
DATETIME_FORMAT_VBA = "%Y-%m-%d %H:%M:%S+00:00"
date_dict = {}
outlook_cal_folder = "[Recurring]"

##################################################
# Functions
##################################################
def is_conflict(curr_start, curr_end, next_start, next_end):
    if curr_start > next_start and curr_start < next_end and curr_end > next_end:
        return True
    elif curr_start <= next_start and curr_end >= next_end:
        return True
    elif curr_start < next_start and curr_end > next_start and curr_end < next_end:
        return True
    
    return False

def insert_dict(pDict, pDate, pStart, pEnd):
    diff = pEnd - pStart
    if pDict.get(pDate) == None:
        pDict[pDate] = diff
    else:
        pDict[pDate] = pDict[pDate] + diff

def increment_date(pDate):
    return datetime.combine(pDate + timedelta(days=1), time.min)

##################################################
# Main
##################################################
# Init Outlook Calendar folder
app = win32com.client.Dispatch("Outlook.Application")
my_namespace = app.GetNamespace("MAPI")
outlook_folder = my_namespace.GetDefaultFolder(9).Folders(outlook_cal_folder) # 9 for Calendar folder

cal_items = outlook_folder.Items
cal_items.IncludeRecurrences = True
cal_items.Sort("[Start]")

prev_start = datetime.min
prev_end = datetime.min
i = 0
for cal in cal_items:
    curr_start_temp = parse_datetime(str(cal.Start), DATETIME_FORMAT_VBA)
    curr_end_temp = parse_datetime(str(cal.End), DATETIME_FORMAT_VBA)
    curr_start = curr_start_temp if (prev_start == datetime.min) else (min(prev_start, curr_start_temp))
    curr_end = curr_end_temp if (prev_end == datetime.min) else (max(prev_end, curr_end_temp))
    next_start = datetime.min
    next_end = datetime.min
    try:
        next_start = parse_datetime(str(cal_items[i + 1].Start), DATETIME_FORMAT_VBA)
        next_end = parse_datetime(str(cal_items[i + 1].End), DATETIME_FORMAT_VBA)
    except IndexError:
        pass

    if next_start != datetime.min and next_end != datetime.min and is_conflict(curr_start, curr_end, next_start, next_end):
        prev_start = curr_start
        prev_end = curr_end
        i += 1
        continue
    elif next_start != datetime.min and next_end != datetime.min and not(is_conflict(curr_start, curr_end, next_start, next_end)):
        prev_start = datetime.min
        prev_end = datetime.min

    if curr_start.date() == curr_end.date():
        insert_dict(date_dict, curr_start.date(), curr_start, curr_end)

    elif curr_start.date() != curr_end.date():
        new_start = increment_date(curr_start.date())
        insert_dict(date_dict, curr_start.date(), curr_start, new_start)
        
        while new_start.date() <= curr_end.date():
            if new_start.date() == curr_end.date():
                insert_dict(date_dict, new_start.date(), new_start, curr_end)
                break
            elif curr_start.date() != curr_end.date():
                curr_start = new_start
                new_start = increment_date(new_start)
                insert_dict(date_dict, curr_start.date(), curr_start, new_start)

    i += 1

for key, value in date_dict.items():
    print(f"{key} : {value.total_seconds() / 3600}")

finalise_app()