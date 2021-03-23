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
my_arg_parser.add_argument("from_date", help="Enter start range in sample format 15/1/1999 3:30 pm")
my_arg_parser.add_argument("to_date", help="Enter end range in sample format 15/1/1999 3:30 pm")
my_arg_parser.add_argument("exclude_prefix", help="Populate list of prefix strings for excluding appointments, delimit with ';'")
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
DATETIME_FORMAT_ARG = "%d/%m/%Y %H:%M %p"
DATETIME_FORMAT_VBA_OUTPUT = "%Y-%m-%d %H:%M:%S+00:00"
CAT_DUE = "Task_Due"
CAT_DO = "Task_Do"
CAT_START = "Task_Start"
folder = r"data\python"
appts = []
appts_all_day = []
date_dict = {}

##################################################
# Functions
##################################################
def vbaDatetimeUtc_to_pyDatetime(pDateTime):
    return datetime.strptime(str(pDateTime), DATETIME_FORMAT_VBA_OUTPUT) + timedelta(hours=8)

def increment_date_to_datetime(pDate):
    return datetime.combine(pDate + timedelta(days=1), time.min)

def decrement_date_to_datetime(pDate):
    return datetime.combine(pDate - timedelta(days=1), time.min)

def is_conflict(pStart_curr, pEnd_curr, pStart_next, pEnd_next):
    if pStart_curr > pStart_next and pStart_curr < pEnd_next and pEnd_curr > pEnd_next:
        return True
    elif pStart_curr <= pStart_next and pEnd_curr >= pEnd_next:
        return True
    elif pStart_curr >= pStart_next and pEnd_curr <= pEnd_next:
        return True
    elif pStart_curr < pStart_next and pEnd_curr > pStart_next and pEnd_curr < pEnd_next:
        return True
    
    return False

def insert_dict_hrs(pDate, pStart, pEnd):
    diff = pEnd - pStart
    if date_dict.get(pDate) == None:
        date_dict[pDate] = MyCls.Day(diff, 0, 0, 0, 0, False)
    else:
        date_dict[pDate].busy_hours += diff

def insert_dict_events(pDate, pCat, pIs_out_of_office):
    if date_dict.get(pDate) == None:
        if pIs_out_of_office:
            date_dict[pDate] = MyCls.Day(timedelta(), 0, 0, 0, 0, pIs_out_of_office)
        elif not(pIs_out_of_office):
            if pCat == CAT_DUE:
                date_dict[pDate] = MyCls.Day(timedelta(), 0, 1, 0, 0, False)
            elif pCat == CAT_DO:
                date_dict[pDate] = MyCls.Day(timedelta(), 0, 0, 1, 0, False)
            elif pCat == CAT_START:
                date_dict[pDate] = MyCls.Day(timedelta(), 0, 0, 0, 1, False)
            else:
                date_dict[pDate] = MyCls.Day(timedelta(), 1, 0, 0, 0, False)
                
    else:
        if pIs_out_of_office:
            date_dict[pDate].is_out_of_office = True
        elif not(pIs_out_of_office):
            if pCat == CAT_DUE:
                date_dict[pDate].due += 1
            elif pCat == CAT_DO:
                date_dict[pDate].do += 1
            elif pCat == CAT_START:
                date_dict[pDate].start += 1
            else:
                date_dict[pDate].all_day_events += 1

def calculate_hrs(pStart, pEnd):
    if pStart.date() == pEnd.date():
        insert_dict_hrs(pStart.date(), pStart, pEnd)
        return

    new_start = increment_date_to_datetime(pStart.date())
    insert_dict_hrs(pStart.date(), pStart, new_start)

    calculate_hrs(new_start, pEnd)

def count_all_days(pStart_date, pEnd_date, pCat, pIs_out_of_office):
    insert_dict_events(pStart_date, pCat, pIs_out_of_office)

    if pStart_date == pEnd_date:
        return

    new_start_date = increment_date_to_datetime(pStart_date).date()

    count_all_days(new_start_date, pEnd_date, pCat, pIs_out_of_office)

##################################################
# Main
##################################################
# Create output folder if not exists
os.makedirs(folder, exist_ok=True)

# Process command arguments
from_date = args.from_date
to_date = args.to_date
exclude_prefixes = args.exclude_prefix.split(';')
exclude_prefixes.pop()

# Init Outlook
app = win32com.client.Dispatch("Outlook.Application")
my_namespace = app.GetNamespace("MAPI")

# Init Calendar folder
outlook_cal_folder = my_namespace.GetDefaultFolder(9) # 9 for Calendar folder

# Get Calendar sub-folders
outlook_cal_folders = []
for cFolder in outlook_cal_folder.Folders:
    outlook_cal_folders.append(cFolder.Name)

# Consolidate items from all calendar folders into 2 lists
fol_i = -1
while (fol_i < len(outlook_cal_folders)):
    # Get calendar appointment items
    cal_items = outlook_cal_folder.Items
    cal_items.Sort("[Start]")
    cal_items.IncludeRecurrences = True
    cal_items_filtered = cal_items.Restrict(f"[Start] >= '{from_date}' and [Start] <= '{to_date}'")

    # Process & filter calendar appointment items
    appt_count = 0
    appt_all_day_count = 0
    for cal in cal_items_filtered:
        if cal.Subject.startswith(tuple(exclude_prefixes)):
            continue

        cStart = vbaDatetimeUtc_to_pyDatetime(cal.StartUTC)
        cEnd = vbaDatetimeUtc_to_pyDatetime(cal.EndUTC)
        if not(cal.AllDayEvent):
            appts.append(MyCls.Appointment(cStart, cEnd, False, "", False))
            appt_count += 1
        elif cal.AllDayEvent:
            cIs_out_of_office = True if (cal.BusyStatus == 3) else False
            appts_all_day.append(MyCls.Appointment(cStart, cEnd, True, cal.Categories, cIs_out_of_office))
            appt_all_day_count += 1

    logger.info(f"Extracted {appt_count} appointments & {appt_all_day_count} all days from {outlook_cal_folder.Name}")
    
    fol_i += 1
    if fol_i >= len(outlook_cal_folders):
        break
    outlook_cal_folder = my_namespace.GetDefaultFolder(9).Folders(outlook_cal_folders[fol_i]) # 9 for Calendar folder

# Sort lists
appts.sort(key=attrgetter("start"))
appts_all_day.sort(key=attrgetter("start"))

# Loop and calculate appointments
prev_start = datetime.min
prev_end = datetime.min
for x in range(len(appts)):
    # Get current date for later use
    curr_start_temp = appts[x].start
    curr_end_temp = appts[x].end
    
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

    # Check if next appointment exists (is current appointment last)
    if next_start != datetime.min and next_end != datetime.min:
        # Skip if conflict with next timed appointment + update prev start & end
        if is_conflict(curr_start, curr_end, next_start, next_end):
            prev_start = curr_start
            prev_end = curr_end
            continue
        # Reset prev start & end if no conflict
        else:
            prev_start = datetime.min
            prev_end = datetime.min

    # Process accumulated hours range
    calculate_hrs(curr_start, curr_end)

# Loop and count all days
for appt in appts_all_day:
    curr_start_date = appt.start.date()
    curr_end_date = decrement_date_to_datetime(appt.end.date()).date() # Decrement end date for comparison
    
    count_all_days(curr_start_date, curr_end_date, appt.cat, appt.is_out_of_office)

# Write date dictionary to txt
with open(f"{folder}/calendar_cal.txt", 'w') as writer:
    writer.write("Day,BusyHours,AllDayEvents,Due,Do,Start,OutOfOffice\n")
    cDate = datetime.strptime(from_date, DATETIME_FORMAT_ARG).date()
    while cDate < datetime.strptime(to_date, DATETIME_FORMAT_ARG).date():
        if date_dict.get(cDate) == None:
            writer.write(f"{cDate},0.0,0,0,0,0,False\n")
        else:
            writer.write(f"{cDate},{date_dict[cDate]}\n")
        cDate += timedelta(days=1)
logger.info(f"Output written to {CURR_DIR}\{folder}")

finalise_app()