# Import from packages
import os
from sys import stdout, exit
import logging
from urllib.request import urlopen
from datetime import datetime
from time import sleep

##################################################
# Variables
##################################################
logger = logging.getLogger("my_logger")
LOG_END = "\n-------------------------"

##################################################
# Functions
##################################################
def initialise_app(project_name, log_level="INFO"):
    # Create directories
    os.makedirs("data", exist_ok=True)
    os.makedirs("data/app", exist_ok=True)

    # Logger
    logger = logging.getLogger("my_logger")
    if log_level == "DEBUG":
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)
    my_log_format = logging.Formatter("[%(asctime)s] %(message)s")
    my_log_stream = logging.StreamHandler(stdout)
    my_log_stream.setFormatter(my_log_format)
    logger.addHandler(my_log_stream)
    my_log_file = logging.FileHandler("data/app/" + project_name + ".log")
    my_log_file.setFormatter(my_log_format)
    logger.addHandler(my_log_file)

def finalise_app(log_message=""):
    if log_message != "":
        logger.info(log_message)
    logger.info("App complete running successfully" + LOG_END)
    exit()

def handle_exception(log_message):
    logger.error(log_message, exc_info=True)
    logger.error("App run fail!" + LOG_END)
    exit()

def is_internet_connected():
    try:
        urlopen("https://www.google.com.sg/", timeout=3)
        return True
    except Exception:
        return False

def parse_datetime(dt_input, input_format="", output_format=""):
    dt = dt_input
    if type(dt_input) == str:
        dt = datetime.strptime(dt_input, input_format)
    
    if output_format != "":
        return dt.strftime(output_format)
    return dt

def str_to_bool(bool_str):
    if bool_str == "True":
        return True
    return False

def wait_file_exist(file_path, my_timeout):
    for i in range(my_timeout):
        if os.path.isfile(file_path):
            return True
        sleep(1)
    return False

def del_file(file_path):
    try:
        os.remove(file_path)
    except FileNotFoundError:
        logger.error("File not found!")