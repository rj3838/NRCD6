from pywinauto.application import Application
import glob
import logging
import os
import sys
import time
# import tkinter
# from tkinter import filedialog
import pywinauto
import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# select the initial (current) file to match

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
initial_file: str = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(initial_file)

# Function to find all the files for that authority


def authority_file_search(file_path_variable: str):

    # get a list of the files matching the Authority of the selected file

    file_search_variable: str = file_path_variable.replace('/', '\\')

    # replace the year from the filename for a generic search [1-3][0-9]{3}
    # file_search_variable: str = re.match(r'.*20[0-9]{2}', file_search_variable)
    # year_str: str = re.match(r'20[0-9]{2}', file_search_variable)
    year_str: str = (re.findall(r'20[0-9]{2}', file_search_variable))[-1]
    file_search_variable = file_search_variable.replace(year_str, "*")
    year_num: int = int(year_str)
    year_num = year_num - 1

    # do a generic search to return a list of files and reverse it

    previous_files_to_match = glob.glob(file_search_variable)

    # remove the first file as it is the initial file (highest year) and reverse it
    # as we want the most recent year first on the list
    previous_files_to_match = previous_files_to_match[:-1]
    previous_files_to_match.reverse()

    return previous_files_to_match


files_to_process: list = authority_file_search(initial_file)



# for each file in the list of previous files
for match_file in files_to_process:



    # match initial selected file against the previous

    # record match numbers

