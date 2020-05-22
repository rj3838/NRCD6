# from pywinauto.application import Application
# import glob
import logging
import os
# import sys
# import time

import tkinter


from tkinter import *
from tkinter import filedialog

# import pandas as pd

sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
root = tkinter.Tk()
root.withdraw()

print('Start')

logger = logging.getLogger('autoNRCD')
logger.setLevel(logging.INFO)

# create console handler, set the file nmae and set level to debug
ch = logging.StreamHandler()
# ch = logging.FileHandler('NRCD.log')

ch.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# handler = logging.FileHandler('info.log')

# add formatter to console handler
ch.setFormatter(formatter)

# add console handler to logger
logger.addHandler(ch)

curr_dir = os.getcwd()

logger.info('Current directory ' + curr_dir)
print('ask directory')


def fun_directory_selector(request_string: str, selected_directory_list: list, search_directory):

    directory_path_string = filedialog.askdirectory(initialdir=search_directory, title=request_string)

    if len(directory_path_string) > 0:
        selected_directory_list.append(directory_path_string)
        fun_directory_selector('Select the next Local Authority Directory or Cancel to end', selected_directory_list,
                               os.path.dirname(directory_path_string))

    return selected_directory_list


root_win = tkinter.Tk()
root.withdraw()

directories_to_process = list()

directories_to_process = fun_directory_selector('Please select a directory containing the data',
                                                directories_to_process,
                                                curr_dir)

print(directories_to_process,sep='\n')
print("\n",directories_to_process)

if len(directories_to_process[0]) > 0:
    logger.info('Working with ')

    for item in directories_to_process:
        logger.info('Working with ' + item)



sys.exit()