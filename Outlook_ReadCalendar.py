# Import from packages
import os
import argparse
import logging
from datetime import datetime, timedelta, time
import win32com.client

# Import from modules
from MyMod import initialise_app, finalise_app, handle_exception, parse_datetime
import MyCls

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
outlook_cal_folder = "[Import]"

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

def insert_dict_hrs(pDict, pDate, pStart, pEnd):
    diff = pEnd - pStart
    if pDict.get(pDate) == None:
        pDict[pDate] = MyCls.Day(diff, 0)
    else:
        pDict[pDate].busy_hours += diff

def insert_dict_events(pDict, pDate):
    if pDict.get(pDate) == None:
        pDict[pDate] = MyCls.Day(timedelta(), 1)
    else:
        pDict[pDate].all_day_events += 1

def increment_date(pDate):
    return datetime.combine(pDate + timedelta(days=1), time.min)

def decrement_date(pDate):
    return datetime.combine(pDate - timedelta(days=1), time.min)

def calculate_hrs(start, end, pDict):
    if start.date() == end.date():
        insert_dict_hrs(pDict, start.date(), start, end)

    elif start.date() != end.date():
        new_start = increment_date(start.date())
        insert_dict_hrs(pDict, start.date(), start, new_start)
        
        while new_start.date() <= end.date():
            if new_start.date() == end.date():
                insert_dict_hrs(pDict, new_start.date(), new_start, end)
                break
            elif new_start.date() != end.date():
                start = new_start
                new_start = increment_date(new_start.date())
                insert_dict_hrs(pDict, start.date(), start, new_start)

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
curr_AllDayEvent = False
i = 0
for cal in cal_items:
    if cal.AllDayEvent:
        curr_AllDayEvent = True

        curr_start_date = parse_datetime(str(cal.Start), DATETIME_FORMAT_VBA).date()
        curr_end_date = decrement_date(parse_datetime(str(cal.End), DATETIME_FORMAT_VBA).date()).date()

        if curr_start_date == curr_end_date:
            insert_dict_events(date_dict, curr_start_date)
        
        elif curr_start_date != curr_end_date:
            insert_dict_events(date_dict, curr_start_date)
            new_start_date = increment_date(curr_start_date).date()
            
            while new_start_date <= curr_end_date:
                if new_start_date == curr_end_date:
                    insert_dict_events(date_dict, new_start_date)
                    break
                elif new_start_date != curr_end_date:
                    insert_dict_events(date_dict, new_start_date)
                    new_start_date = increment_date(new_start_date).date()
        
    elif not(cal.AllDayEvent):
        curr_AllDayEvent = False

        curr_start_temp = parse_datetime(str(cal.Start), DATETIME_FORMAT_VBA)
        curr_end_temp = parse_datetime(str(cal.End), DATETIME_FORMAT_VBA)
        
        if not(is_conflict(prev_start, prev_end, curr_start_temp, curr_end_temp)) and prev_start != datetime.min and prev_end != datetime.min:
            calculate_hrs(prev_start, prev_end, date_dict)
            prev_start = datetime.min
            prev_end = datetime.min

        curr_start = curr_start_temp if (prev_start == datetime.min) else (min(prev_start, curr_start_temp))
        curr_end = curr_end_temp if (prev_end == datetime.min) else (max(prev_end, curr_end_temp))
        next_start = datetime.min
        next_end = datetime.min
        next_AllDayEvent = False
        try:
            next_cal = cal_items[i + 1]
            next_start = parse_datetime(str(next_cal.Start), DATETIME_FORMAT_VBA)
            next_end = parse_datetime(str(next_cal.End), DATETIME_FORMAT_VBA)
            next_AllDayEvent = next_cal.AllDayEvent
        except IndexError:
            pass

        if next_AllDayEvent:
            prev_start = curr_start
            prev_end = curr_end
            i += 1
            continue
        
        if next_start != datetime.min and next_end != datetime.min:
            if is_conflict(curr_start, curr_end, next_start, next_end):
                prev_start = curr_start
                prev_end = curr_end
                i += 1
                continue
            else:
                prev_start = datetime.min
                prev_end = datetime.min

        calculate_hrs(curr_start, curr_end, date_dict)
    
    i += 1

if curr_AllDayEvent and prev_start != datetime.min and prev_end != datetime.min:
    calculate_hrs(prev_start, prev_end, date_dict)

for key, value in date_dict.items():
    print(f"{key} : {value}")

finalise_app()