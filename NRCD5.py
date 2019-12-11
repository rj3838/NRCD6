
#import pywinauto
#from pywinauto.application import Application
import glob
import time
import pandas as pd
import logging
import sys
from tkinter import filedialog
import os

import tkinter

sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
root = tkinter.Tk()
root.withdraw()

logger = logging.getLogger('autoNRCD')
logger.setLevel(logging.INFO)

# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# add formatter to ch
ch.setFormatter(formatter)

# add ch to logger
logger.addHandler(ch)


#def search_for_file_path():
#import tkinter.filedialog
curr_dir = os.getcwd()
#print(curr_dir)
logger.info('Current directory' + curr_dir)

file_path_variable = filedialog.askdirectory(initialdir=curr_dir,
                                              title='Please select a directory containing the data')
if len(file_path_variable) > 0:
        # print("You chose: %s" % tempdir)
    logger.info('Working with' + file_path_variable)
#return tempdir


def nrcd_running_check(number_of_windows : str):
    from pywinauto import Application

    tb = Application(backend="uia").connect(title='Taskbar')  # backend is important!!!

    running_apps = tb.Taskbar.child_window(title="Running applications", control_type="ToolBar")

    nrcd_windows = "NRCD - " + number_of_windows + "running windows"
    print(nrcd_windows)

    if nrcd_windows in [w.window_text() for w in running_apps.children()]:
        wait_flag = True
    else:
        wait_flag = False

    return wait_flag


#file_path_variable = search_for_file_path()
# file_path_variable = tempdir
print("\nfile_path_variable = ", file_path_variable)

#app = Application.start(path="C:/Program Files (x86)/NRCD/NRCD.exe",backend="uia")
from pywinauto.application import Application

app = Application(backend="uia").start('C:/Program Files (x86)/NRCD/NRCD.exe')

app.window(best_match='National Roads Condition Database',
           top_level_only=True).child_window(best_match='SCANNER').click()

time.sleep(15)

app.window(best_match='',top_level_only=True) \
    .print_control_identifiers()

app.window(best_match='',top_level_only=True) \
    .child_window(best_match='Enter System').click()
# app.window(best_match='National Roads Condition Database',
#               top_level_only=True).child_window(best_match='Loading').click()

# from pathlib import Path


# Loading section

def dataloading(app, file_path_variable):

    app.window(best_match='National Roads Condition Database - Version*', top_level_only=True).child_window(
        best_match='Loading').click()

# if the file_path_variable directory contains a file 'BatchList' use 'Select Batch File'

    if os.path.isfile(file_path_variable + '/BatchList.txt'):
        filename = file_path_variable + '/BatchList.txt'
        print("\nfile name exists using Select", filename)

        time.sleep(15)

        app.window(best_match='National Roads Condition Database Version', top_level_only=True) \
            .child_window(best_match='Select Batch file').click()

        filename = filename.replace('/', '\\')
        print("\nfile exists", filename)
        time.sleep(15)
        app2 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
        print("\nConnect app2 filename is ", filename)
        app2.window(title_re='Select a batch file').type_keys(filename, with_spaces=True)
        app2.window(title_re='Select a batch file').type_keys('%o')
        del app2
    else:
        # else pick the first .hmd file and use 'Create Batch File'
        print("\nfile missing")

        # app.window(best_match='National Roads Condition Database - Version', top_level_only=True).child_window(
        #   best_match='OK').click()

        file_search_variable = (file_path_variable + '/*.hmd').replace('/', '\\')
        print("\nfile_search_variable = ", file_search_variable)
        filename = glob.glob(file_search_variable)
        # filename = filename[0]
        print("\nFile found : ", filename[0])
        # time.sleep(30)
        app.window(best_match='National Roads Condition Database Version *', top_level_only=True) \
            .child_window(best_match='Create Batch file').click()

        time.sleep(15)

        app.Dialog.OK.click()

        time.sleep(15)
        app3 = Application(backend="uia").connect(title='Create a file in the required directory')
        print("\nconnect app3")
        time.sleep(15)
        app3.window(title_re='Create a file in the required directory').type_keys(filename[0], with_spaces=True)
        app3.window(title_re='Create a file in the required directory').type_keys('%o')
        del app3

    # if the file_path_variable directory string contains WDM
    if "WDM" in file_path_variable:
        # then Survey Contractor = WDM
        surveycontractor = "WDM"

        # else if the directory string contains 'G-L' select survey contractor 'Ginger-Lehmann'
    elif "G-L" in file_path_variable:
        surveycontractor = "Ginger-Lehmann"
    else:
        surveycontractor = "Unknown"

    print(surveycontractor)

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Survey Contractor').Edit.type_keys(surveycontractor)

    # click 'OK' to close the data loading window
    app.window(best_match='National Roads Condition Database').child_window(best_match='OK').click()

    # back on the main screen, click the process radio button then the actual 'Process' button
    # app.window(best_match='National Roads Condition Database - Version', top_level_only=True, auto_id=18).click()
    app.window(best_match='National Roads Condition Database - Version',
               top_level_only=True).child_window(best_match='Process').click()

    logger.info('Starting loading for ' + filename[0])
    # wait for the loading to finish

    while nrcd_running_check("2"):
        logger.info('waiting for loading to finish')
        time.sleep(60)

    return

# First attributes section (this sets the local authority)


def assign_la(app, file_path_variable):

    logger.info('starting local authority assignment')

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Attributes').click()
    time.sleep(15)

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).print_control_identifiers()

    #.child_window(title="National Roads Condition Database", control_type="Window") \
    time.sleep(15)

    groupcontrol = app.window(best_match='National Roads Condition Database - Version *') \
       .child_window(title="Local Authority Attribute", auto_id="3", control_type="Group")

    groupcontrol.child_window(title="Select Batch File", auto_id="7", control_type="Button") \
        .click()

    # app.window(best_match='National Roads Condition Database - Version *') \
    #  .child_window(title="Local Authority Attribute", auto_id="3", control_type="Group") \
    #   .child_window(title="Select Batch File", auto_id="7", control_type="Button") \
    #   .click()

    filename = file_path_variable + '/BatchList.txt'
    print("\nfilename is ", filename)
    filename = filename.replace('/', '\\')
    time.sleep(15)
    app4 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
    print("\nConnect app4 filename is ", filename)
    time.sleep(15)
    app4.window(title_re='Select a batch file').type_keys(filename, with_spaces=True)
    app4.window(title_re='Select a batch file').type_keys('%o')
    del app4
    # use the location of the BatchList.txt to extract the local authority.

    local_authority = filename.split("\\")[-2]

    # print(local_authority)

    logger.info('starting to assign attributes for ' + local_authority)
    # Load the human name to database name csv table

    lookup_table = '//trllimited/data/INF_ScannerQA/Audit_Reports/Tools/LA_to_NRCD_Name.csv'

    la_lookup_table = pd.read_csv(lookup_table, index_col='Audit_Report_name', dtype={'INDEX': str})

    la_db_name : str = la_lookup_table.loc[local_authority, "NRCD_Name"]

    survey_year = "2019/20"
    print(la_db_name)
    logger.info('DB name is  ' + la_db_name)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group").print_control_identifiers()
    time.sleep(15)

    # app.window(best_match='National Roads Condition Database - Version *') \
    # .child_window(title="Local Authority Attribute", control_type="Group") \
    # .child_window(best_match='3',control_type="ComboBox").type_keys("%{DOWN}")
    # app.window(best_match='National Roads Condition Database - Version*') \
    # .child_window(best_match='Local Authority Attribute', control_type="Group") \
    # .child_window(auto_id='6', control_type="ComboBox").select(la_db_name)
    print(la_db_name)

    batch_combobox1 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='6', control_type="ComboBox").wait('exists enabled visible ready')

    from pywinauto.controls.win32_controls import ComboBoxWrapper
    ComboBoxWrapper(batch_combobox1).select(la_db_name)

    batch_combobox2 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='4', control_type="ComboBox").wait('exists enabled visible ready')

    ComboBoxWrapper(batch_combobox2).select(" 2019/20")

    time.sleep(5)

    print(survey_year)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Assign Local Authority", auto_id="5", control_type="Button") \
        .click()

    logger.info('waiting for LA assignment to complete')

    while nrcd_running_check("2"):
        logger.info('waiting for loading to finish')
        time.sleep(60)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title='NRCD', control_type="Window") \
        .child_window(title="OK", auto_id="2", control_type="Button").click()

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Exit", auto_id="11", control_type="Button").click()

    return


def fitting(app):

    logger.info('starting fitting')

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Fitting').click()

    time.sleep(5)

    app.window(best_match='National Roads Condition Database', top_level_only=True) \
        .child_window(best_match='Fit using a grid').click()

    app.window(best_match='National Roads Condition Database', top_level_only=True) \
        .child_window(best_match='Fit unfitted data').click()

    batch_combobox4 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Year options', control_type="Group") \
        .child_window(control_type="ComboBox").wait('exists enabled visible ready')

    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper
    ComboBoxWrapper(batch_combobox4).select(" 2019/20")

    time.sleep(5)

    main_screen = app.window(best_match='National Roads Condition Database - V*')#\
    #    .child_window(best_match='National Roads Condition Database')

    main_screen.child_window(title="Exit", auto_id="12", control_type="Button").click()

    #main_screen_group = main_screen.child_window(auto_id="11", control_type="Group")

    main_screen_ProcessCheckbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_ProcessCheckbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_ProcessCheckbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_ProcessCheckbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    ButtonWrapper(main_screen_ProcessCheckbox).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox2).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox3).check_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox4).uncheck_by_click()

    main_screen.Process.click()

    time.sleep(60)

    while nrcd_running_check("2"):
        logger.info('waiting for Fitting to complete')
        time.sleep(60)

    logger.info('fitting complete')
    return


# dataloading(app, file_path_variable)

# assign_la(app, file_path_variable)

fitting(app)

# attributes

# scanner_qa

print('End')
