# Import from packages
import os
import argparse
import logging
from datetime import datetime, timedelta, time
from operator import attrgetter
import win32com.client

# Import from modules
from MyMod import initialise_app, finalise_app, handle_exception
import MyCls

# Initialise project
CURR_DIR, CURR_FILE = os.path.split(__file__)
PROJ_NAME = CURR_FILE.split('.')[0]

# Get command line arguments
my_arg_parser = argparse.ArgumentParser(description=f"{PROJ_NAME}")
my_arg_parser.add_argument("startdate", help="Enter start date in %Y-%m-%d format")
my_arg_parser.add_argument("enddate", help="Enter start date in %Y-%m-%d format")
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
DATETIME_FORMAT_ARG = "%Y-%m-%d"
DATETIME_FORMAT_VBA_OUTPUT = "%Y-%m-%d %H:%M:%S+00:00"
CAT_DUE = "Task_Due"
CAT_DO = "Task_Do"
CAT_START = "Task_Start"
folder = r"data\python"
appts = []
date_dict = {}

##################################################
# Functions
##################################################
def is_conflict(curr_start, curr_end, next_start, next_end):
    if curr_start > next_start and curr_start < next_end and curr_end > next_end:
        return True
    elif curr_start <= next_start and curr_end >= next_end:
        return True
    elif curr_start >= next_start and curr_end <= next_end:
        return True
    elif curr_start < next_start and curr_end > next_start and curr_end < next_end:
        return True
    
    return False

def insert_dict_hrs(pDict, pDate, pStart, pEnd):
    diff = pEnd - pStart
    if pDict.get(pDate) == None:
        pDict[pDate] = MyCls.Day(diff, 0, 0, 0, 0)
    else:
        pDict[pDate].busy_hours += diff

def insert_dict_events(pDict, pDate, pCat):
    if pDict.get(pDate) == None:
        if pCat == CAT_DUE:
            pDict[pDate] = MyCls.Day(timedelta(), 0, 1, 0, 0)
        elif pCat == CAT_DO:
            pDict[pDate] = MyCls.Day(timedelta(), 0, 0, 1, 0)
        elif pCat == CAT_START:
            pDict[pDate] = MyCls.Day(timedelta(), 0, 0, 0, 1)
        else:
            pDict[pDate] = MyCls.Day(timedelta(), 1, 0, 0, 0)
    else:
        if pCat == CAT_DUE:
            pDict[pDate].due += 1
        elif pCat == CAT_DO:
            pDict[pDate].do += 1
        elif pCat == CAT_START:
            pDict[pDate].start += 1
        else:
            pDict[pDate].all_day_events += 1

def process_arg_date(date_str, day_delta):
    return datetime.strptime(date_str, DATETIME_FORMAT_ARG) + timedelta(days=day_delta)

def vbaDatetimeUtc_to_pyDatetime(pDateTime):
    return datetime.strptime(str(pDateTime), DATETIME_FORMAT_VBA_OUTPUT) + timedelta(hours=8)

def increment_date_to_datetime(pDate):
    return datetime.combine(pDate + timedelta(days=1), time.min)

def decrement_date_to_datetime(pDate):
    return datetime.combine(pDate - timedelta(days=1), time.min)

def calculate_hrs(start, end, pDict):
    if start.date() == end.date():
        insert_dict_hrs(pDict, start.date(), start, end)

    elif start.date() != end.date():
        new_start = increment_date_to_datetime(start.date())
        insert_dict_hrs(pDict, start.date(), start, new_start)
        
        while new_start.date() <= end.date():
            if new_start.date() == end.date():
                insert_dict_hrs(pDict, new_start.date(), new_start, end)
                break
            elif new_start.date() != end.date():
                start = new_start
                new_start = increment_date_to_datetime(new_start.date())
                insert_dict_hrs(pDict, start.date(), start, new_start)

##################################################
# Main
##################################################
# Create output folder if not exists
os.makedirs(folder, exist_ok=True)

# Process dates command arguments
start_date = process_arg_date(args.startdate, 0)
end_date = process_arg_date(args.enddate, 1)

# Init Outlook
app = win32com.client.Dispatch("Outlook.Application")
my_namespace = app.GetNamespace("MAPI")

# Init Calendar folder
outlook_cal_folder = my_namespace.GetDefaultFolder(9) # 9 for Calendar folder

# Get Calendar sub-folders
outlook_cal_folders = []
for c_folder in outlook_cal_folder.Folders:
    outlook_cal_folders.append(c_folder.Name)

# Do-while
fol_i = -1
while (fol_i < len(outlook_cal_folders)):
    # Get calendar appointment items
    cal_items = outlook_cal_folder.Items
    cal_items.IncludeRecurrences = True

    # Filter calendar appointment items by dates
    appt_count = 0
    for cal in cal_items:
        c_start = vbaDatetimeUtc_to_pyDatetime(cal.StartUTC)
        c_end = vbaDatetimeUtc_to_pyDatetime(cal.EndUTC)
        if c_start >= start_date and c_end <= end_date:
            appts.append(MyCls.Appointment(c_start, c_end, cal.AllDayEvent, cal.Categories))
            appt_count += 1

    logger.info(f"Extracted {appt_count} appointments from {outlook_cal_folder.Name}")
    
    fol_i += 1
    if fol_i >= len(outlook_cal_folders):
        break
    outlook_cal_folder = my_namespace.GetDefaultFolder(9).Folders(outlook_cal_folders[fol_i]) # 9 for Calendar folder

# Sort appts list
appts.sort(key=attrgetter("start"))

# Loop and process appointments
prev_start = datetime.min
prev_end = datetime.min
curr_AllDayEvent = False
for x in range(len(appts)):
    # Handle AllDayEvent appointments
    if appts[x].is_all_day:
        curr_AllDayEvent = True

        curr_start_date = appts[x].start.date()
        curr_end_date = decrement_date_to_datetime(appts[x].end.date()).date() # Decrement end date for comparison

        # Update for current date
        insert_dict_events(date_dict, curr_start_date, appts[x].cat)
        
        # Loop subsequent dates
        if curr_start_date != curr_end_date:
            new_start_date = increment_date_to_datetime(curr_start_date).date()
            
            while new_start_date <= curr_end_date:
                insert_dict_events(date_dict, new_start_date, appts[x].cat)
                if new_start_date == curr_end_date:
                    break
                elif new_start_date != curr_end_date:
                    new_start_date = increment_date_to_datetime(new_start_date).date()
        
    # Handle timed appointments
    elif not(appts[x].is_all_day):
        curr_AllDayEvent = False

        # Get current date for later use
        curr_start_temp = appts[x].start
        curr_end_temp = appts[x].end
        
        # Clear previous pending due to next_AllDayEvent skip
        if not(is_conflict(prev_start, prev_end, curr_start_temp, curr_end_temp)) and prev_start != datetime.min and prev_end != datetime.min:
            calculate_hrs(prev_start, prev_end, date_dict)
            prev_start = datetime.min
            prev_end = datetime.min

        # Initialise current and next appointment details
        curr_start = curr_start_temp if (prev_start == datetime.min) else (min(prev_start, curr_start_temp))
        curr_end = curr_end_temp if (prev_end == datetime.min) else (max(prev_end, curr_end_temp))
        next_start = datetime.min
        next_end = datetime.min
        next_AllDayEvent = False
        if x + 1 < len(appts):
            next_appt = appts[x + 1]
            next_start = next_appt.start
            next_end = next_appt.end
            next_AllDayEvent = next_appt.is_all_day

        # Skip if next appointment is AllDayEvent (for conflict management with subsequent timed events) + update prev start end
        if next_AllDayEvent:
            prev_start = curr_start
            prev_end = curr_end
            continue
        
        # Check if next appointment exists (is current appointment last)
        if next_start != datetime.min and next_end != datetime.min:
            # Skip if conflict with next timed appointment + update prev start end
            if is_conflict(curr_start, curr_end, next_start, next_end):
                prev_start = curr_start
                prev_end = curr_end
                continue
            # Reset prev start end if no conflict
            else:
                prev_start = datetime.min
                prev_end = datetime.min

        # Process accumulated hours range
        calculate_hrs(curr_start, curr_end, date_dict)

# Handle skip if last appointment is AllDayEvent
if curr_AllDayEvent and prev_start != datetime.min and prev_end != datetime.min:
    calculate_hrs(prev_start, prev_end, date_dict)

# Write date dictionary to txt
with open(f"{folder}/calendar_cal.txt", 'w') as writer:
    writer.write("Day,BusyHours,AllDayEvents,Due,Do,Start\n")
    for key, value in date_dict.items():
        writer.write(f"{key},{value}\n")
logger.info(f"Output written to {CURR_DIR}\{folder}")

finalise_app()