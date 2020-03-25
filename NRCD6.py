#from pywinauto.application import Application
import glob
import logging
import os
import sys
import time
import tkinter
from tkinter import filedialog

import pandas as pd

sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
root = tkinter.Tk()
root.withdraw()

print('Start')

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

curr_dir = os.getcwd()

logger.info('Current directory ' + curr_dir)
print('ask directory')
file_path_variable = filedialog.askdirectory(initialdir=curr_dir, title='Please select a directory containing the data')

if len(file_path_variable) > 0:
    logger.info('Working with ' + file_path_variable)


def nrcd_running_check(number_of_windows: str):
    from pywinauto import Application

    tb = Application(backend="uia").connect(title='Taskbar')  # backend is important!!!

    running_apps = tb.Taskbar.child_window(title="Running applications", control_type="ToolBar")

    nrcd_windows = "NRCD.exe - " + number_of_windows + " running windows"
    print(nrcd_windows)

    # print([w.window_text() for w in running_apps.children()])

    if nrcd_windows in [w.window_text() for w in running_apps.children()]:
        wait_flag = True
    else:
        wait_flag = False

    return wait_flag


print("\nfile_path_variable = ", file_path_variable)

from pywinauto.application import Application

app = Application(backend="uia").start('C:/Users/rjaques/Software/NRCD/Current Version/NRCD.exe')

app.window(best_match='National Roads Condition Database',
           top_level_only=True).child_window(best_match='SCANNER').click()

time.sleep(5)

# app.window(best_match='', top_level_only=True) \
# .print_control_identifiers()

app.window(best_match='', top_level_only=True) \
    .child_window(best_match='Enter System').click()

time.sleep(30)


# Loading section


def dataloading(app, file_path_variable):
    app.window(best_match='National Roads Condition Database - Version*', top_level_only=True).child_window(
        best_match='Loading').click()

    # if the file_path_variable directory contains a file 'BatchList' use 'Select Batch File'

    if os.path.isfile(file_path_variable + '/BatchList.txt'):
        filename = file_path_variable + '/BatchList.txt'
        print("\nfile name exists using Select", filename)

        time.sleep(30)

        app.window(best_match='National Roads Condition Database Version', top_level_only=True) \
            .child_window(best_match='Select Batch file').click()

        filename = filename.replace('/', '\\')
        print("\nfile exists", filename)
        time.sleep(15)
        app2 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
        print("\nConnect app2 filename is ", filename)
        # edit_text_box1 = app2.window(title_re='Select a batch file') \
        #    .child_window(best_match="File name:")
        # from pywinauto.controls.win32_controls import EditWrapper
        # EditWrapper(edit_text_box1).set_text(filename)
        # app2.window(title_re='Select a batch file').type_keys(filename, with_spaces=True)
        app2.window(title_re='Select a batch file').File_name_Edit.set_text(filename)
        app2.window(title_re='Select a batch file').print_control_identifiers()
        # app2.window(title_re='Select a batch file').type_keys('%o')

        batch_splitbutton2 = app2.window(title_re='Select a batch file') \
            .child_window(auto_id='1', control_type="SplitButton")
        from pywinauto.controls.win32_controls import ButtonWrapper
        ButtonWrapper(batch_splitbutton2).click()

        # app2.window(title_re='Select a batch file').OpenSplitButton.click_input
        # app2.window(title_re='Select a batch file') \
        #   .child_window(title='Pane1') \
        #   .child_window(title='Open', auto_id=1).click()
        #    .child_window(title='Open', auto_id=1, control_type="UIA_SplitButtonControlTypeId").click()
        # .child_window(title='SplitButton6').click()

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

        time.sleep(30)

        app.window(best_match='National Roads Condition Database Version *') \
            .child_window(best_match='NRCD') \
            .child_window(best_match='OK').click()

        # app.window(best_match='National Roads Condition Database Version *') \
        #    .child_window(best_match='OK').click()

        time.sleep(30)
        app3 = Application(backend="uia").connect(title='Create a file in the required directory')
        print("\nconnect app3")
        time.sleep(15)
        # edit_text_box2 = app3.window(title_re='Select a batch file') \
        #    .child_window(best_match="File name:")
        # from pywinauto.controls.win32_controls import EditWrapper
        # EditWrapper(edit_text_box2).set_text(filename)
        app3.window(title_re='Create a file in the required directory') \
            .File_name_Edit.set_text(filename[0])
        # app3.window(title_re='Create a file in the required directory').type_keys(filename[0], with_spaces=True)
        app3.window(title_re='Create a file in the required directory').print_control_identifiers()
        # app3.window(title_re='Create a file in the required directory') \
        # .Open3_SplitButton.click()
        batch_splitbutton1 = app3.window(title_re='Create a file in the required directory') \
            .child_window(auto_id='1', control_type="SplitButton")
        from pywinauto.controls.win32_controls import ButtonWrapper
        ButtonWrapper(batch_splitbutton1).click()
        # child_window(title="Open", auto_id="1", control_type="SplitButton")
        # .child_window(best_match='Open3').click()
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
    app.window(best_match='National Roads Condition Database') \
        .child_window(best_match='OK').click()

    # back on the main screen, click the process radio button then the actual 'Process' button

    app.window(best_match='National Roads Condition Database - Version *',
               top_level_only=True).child_window(best_match='Process').click()

    print(filename)

    logger.info('Starting loading with ' + filename[0])

    time.sleep(60)

    # wait for the loading to finish. It checks the number of windows open for NECD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check("2"):
        logger.info('waiting for loading to complete')
        time.sleep(120)

    logger.info('loading completed')

    return

# First attributes section (this sets the local authority)


def assign_la(app, file_path_variable):
    logger.info('starting local authority assignment')

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Attributes').click()
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
    logger.info('File name is ' + filename)
    filename = filename.replace('/', '\\')
    time.sleep(15)
    app4 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
    print("\nConnect app4 filename is ", filename)
    logger.info('Connecting to the batch file selection with ' + filename)
    time.sleep(60)
    # app4.window(title_re='Select a batch file').File_name_Edit.set_text(filename)
    # app4.window(title_re='Select a batch file').type_keys(filename, with_spaces=True)
    app4.window(title_re='Select a batch file').print_control_identifiers()
    app4.window(title_re='Select a batch file').File_name_Edit.set_text(filename)
    app4.window(title_re='Select a batch file').print_control_identifiers()

    # app4.window(title_re='Select a batch file').type_keys('%o')
    batch_splitbutton1 = app4.window(title_re='Select a batch file') \
        .child_window(auto_id='1', control_type="SplitButton")
    from pywinauto.controls.win32_controls import ButtonWrapper
    ButtonWrapper(batch_splitbutton1).click()
    del app4

    # use the location of the BatchList.txt to extract the local authority. It's two from the end.

    local_authority = filename.split("\\")[-2]

    # print(local_authority)

    logger.info('starting to assign attributes for ' + local_authority)

    # Load the human name to database name csv table

    lookup_table = '//trllimited/data/INF_ScannerQA/Audit_Reports/Tools/LA_to_NRCD_Name.csv'

    la_lookup_table = pd.read_csv(lookup_table, index_col='Audit_Report_name', dtype={'INDEX': str})

    la_db_name: str = la_lookup_table.loc[local_authority, "NRCD_Name"]

    survey_year = "2019/20"
    # print(la_db_name)
    logger.info('DB name for the LA is ' + la_db_name)

    # app.window(best_match='National Roads Condition Database - Version*') \
    #   .child_window(best_match='Local Authority Attribute', control_type="Group").print_control_identifiers()
    time.sleep(15)

    # app.window(best_match='National Roads Condition Database - Version *') \
    # .child_window(title="Local Authority Attribute", control_type="Group") \
    # .child_window(best_match='3',control_type="ComboBox").type_keys("%{DOWN}")
    # app.window(best_match='National Roads Condition Database - Version*') \
    # .child_window(best_match='Local Authority Attribute', control_type="Group") \
    # .child_window(auto_id='6', control_type="ComboBox").select(la_db_name)
    # print(la_db_name)

    batch_combobox1 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='6', control_type="ComboBox").wait('exists enabled visible ready')

    from pywinauto.controls.win32_controls import ComboBoxWrapper
    ComboBoxWrapper(batch_combobox1).select(la_db_name)

    batch_combobox2 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='4', control_type="ComboBox")  # .wait('exists enabled visible ready')

    ComboBoxWrapper(batch_combobox2).select(" 2019/20")

    time.sleep(15)

    # print(survey_year)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Assign Local Authority", auto_id="5", control_type="Button") \
        .click()

    logger.info('waiting for LA assignment to complete')

#  the following contains 10000 seconds. This is to stop the wait timing out, it retries each 90 secs
#  but the attributes should be finished in under 6 hours... or even one hour.
    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title='NRCD', control_type="Window") \
        .child_window(title="OK", auto_id="2", control_type="Button") \
        .wait("exists ready", timeout=10000, retry_interval=90)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title='NRCD', control_type="Window") \
        .child_window(title="OK", auto_id="2", control_type="Button").click()

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Exit", auto_id="11", control_type="Button").click()

    return


def fitting(app):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper
    # from pywinauto.controls import uia_controls
    # from pywinauto.controls import common_controls

    main_screen = app.window(best_match='National Roads Condition Database - V*')

    main_screen_ProcessCheckbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_ProcessCheckbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_ProcessCheckbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_ProcessCheckbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    ButtonWrapper(main_screen_ProcessCheckbox).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox2).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox3).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox4).uncheck_by_click()

    logger.info('starting fitting')

    time.sleep(30)

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Fitting').click()

    time.sleep(10)

    # set the check boxes for fitting the data using a grid.

    ButtonWrapper(app.window(best_match='National Roads Condition Database')
                  .child_window(title="Fit unfitted data", auto_id="22", control_type="RadioButton")) \
        .check_by_click()

    ButtonWrapper(app.window(best_match='National Roads Condition Database')
                  .child_window(title="Fit using a grid", auto_id="14", control_type="RadioButton")) \
        .check_by_click()

    ButtonWrapper(app.window(best_match='National Roads Condition Database')
                  .child_window(title="SCANNER", auto_id="7", control_type="RadioButton")) \
        .check_by_click()

    # and set the survey year

    batch_combobox24 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Year options:', auto_id="21", control_type="Group") \
        .child_window(auto_id='24', control_type="ComboBox")

    ComboBoxWrapper(batch_combobox24).select(" 2019/20")

    time.sleep(5)

    # with everything set click the OK button to return to the main window.

    main_screen.child_window(title="OK", auto_id="11", control_type="Button").click()

    time.sleep(15)

    # app.window(best_match='National Roads Condition Database - Version*') \
    #    .child_window(best_match='Year options', control_type="Group") \
    #    .child_window(control_type="ComboBox").wait('exists enabled visible ready')
    # main_screen_group = main_screen.child_window(auto_id="11", control_type="Group")

    # main_screen_ProcessCheckbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    # main_screen_ProcessCheckbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    # main_screen_ProcessCheckbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    # main_screen_ProcessCheckbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    # ButtonWrapper(main_screen_ProcessCheckbox).uncheck_by_click()
    # ButtonWrapper(main_screen_ProcessCheckbox2).uncheck_by_click()
    # ButtonWrapper(main_screen_ProcessCheckbox3).check_by_click()
    # ButtonWrapper(main_screen_ProcessCheckbox4).uncheck_by_click()

    # the 'process' click box is already set so click on the main scree Process button at the bottom
    # then wait for the fitting to complete.

    main_screen.Process.click()

    time.sleep(60)

    # wait for the loading to finish. It checks the number of windows open for NECD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check("2"):
        logger.info('waiting for Fitting to complete')
        time.sleep(90)

    logger.info('fitting complete')

    # to the main code block.
    return


def assign_urb_rural(app):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper

    main_screen = app.window(best_match='National Roads Condition Database - V*')

    # on the main screen turn all the check boxes off.

    main_screen_ProcessCheckbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_ProcessCheckbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_ProcessCheckbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_ProcessCheckbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    ButtonWrapper(main_screen_ProcessCheckbox).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox2).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox3).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox4).uncheck_by_click()

    logger.info('starting urban & rural attributes')

    time.sleep(15)

    # start the attributes section.

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Attributes').click()

    time.sleep(10)

    # to set the urban rural attribute the selections are all made in the 'Attributes Options' part of the window.
    # wait for it to exist.

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Attributes Options", auto_id="12", control_type="Group") \
        .wait("exists ready", timeout=90, retry_interval=60)

    # set groupcontrol to that part of the window.

    groupcontrol = app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Attributes Options", auto_id="12", control_type="Group")

    # groupcontrol.print_control_identifiers()

    # Select SCANNER and the survey year and click on OK

    ComboBoxWrapper(groupcontrol.child_window(auto_id="14",
                                              control_type="ComboBox")) \
        .select("SCANNER")

    ComboBoxWrapper(groupcontrol.child_window(auto_id="15",
                                              control_type="ComboBox")) \
        .select(" 2019/20")

    app.window(best_match='National Roads Condition Database').child_window(best_match='OK').click()

    # back to the main window where the process click box is already set by the NRCD prog and it's waiting
    # for the process button at the bottom to be clicked.

    main_screen.Process.click()

    time.sleep(60)

    # after an appropriate time start waiting for the processing to finish.

    # wait for the window to close. It checks the number of windows open for NECD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check("2"):
        logger.info('waiting for Urban/Rural Attributes to complete')
        time.sleep(90)

    logger.info('U/R attributes complete')

    return  # to the main code block


def scanner_qa(app, file_path_variable):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper
    import re  # there is a regex search below !

    main_screen = app.window(best_match='National Roads Condition Database - V*')

    # turn all the process check boxes off.

    main_screen_ProcessCheckbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_ProcessCheckbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_ProcessCheckbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_ProcessCheckbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    ButtonWrapper(main_screen_ProcessCheckbox).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox2).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox3).uncheck_by_click()
    ButtonWrapper(main_screen_ProcessCheckbox4).uncheck_by_click()

    logger.info('starting Scanner QA output')

    filename = file_path_variable.replace('/', '\\')
    print(filename)
    # time.sleep(15)

    # use the location of the file_path_variable to extract the local authority, nation and year.

    local_authority = filename.split("\\")[-1]
    nation = filename.split("\\")[-2]
    year = re.search('[1-2][0-9]{3}', filename).group(0)

    print(local_authority)
    print(nation)
    print(year)

    logger.info('starting SCANNER QA for ' + local_authority)

    # Load the human name to database name csv table

    lookup_table_filename = '//trllimited/data/INF_ScannerQA/Audit_Reports/Tools/LA_to_NRCD_Name.csv'

    la_lookup_table = pd.read_csv(lookup_table_filename, index_col='Audit_Report_name', dtype={'INDEX': str})

    la_db_name: str = la_lookup_table.loc[local_authority, "NRCD_Name"]

    # enter the Scanner QA section of NRCD and wait for the Survey QA options group to exist.

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Scanner QA').click()

    time.sleep(10)

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Survey QA Options", auto_id="9", control_type="Group") \
        .wait("exists ready", timeout=90, retry_interval=60)

    groupcontrol = app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Survey QA Options", auto_id="9", control_type="Group")

    #groupcontrol.print_control_identifiers()

    # exclude the previous year and the U roads (uncheck_by_click) then select the LA abd survey year.

    ButtonWrapper(groupcontrol
                  .child_window(title="Include Previous Year", control_type="CheckBox")) \
        .uncheck_by_click()

    ButtonWrapper(groupcontrol
                  .child_window(title="U Roads", control_type="CheckBox")) \
        .uncheck_by_click()

    ComboBoxWrapper(groupcontrol.child_window(auto_id="24",
                                              control_type="ComboBox")) \
        .select(la_db_name)

    ComboBoxWrapper(groupcontrol.child_window(auto_id="25",
                                              control_type="ComboBox")) \
        .select(" 2019/20")

    # Export the data

    groupcontrol.child_window(auto_id="26",
                              title="Export QA Data").click()

    # print("Waiting for a user !")

    # build output file name.

    # output_file_name = os.path.normpath("//trllimited/data/INF_ScannerQA/Audit_Reports/NRCD Data"
    output_file_name = os.path.normpath("C:/Users/rjaques/temp"
                                        + "/" + nation +
                                        "/" + local_authority + "_" + year + ".csv")

    # add the year combination (year and '-' and  2 digit next year so
    # convert year string to numeric, add one, convert back to string and use the last 2 chars

    print(year)
    print(int(year))
    print(int(year) + 1)
    print(str(int(year) + 1))

    next_year = str(int(year) + 1)[-2:]

    print(next_year)

    survey_period = year + '-' + next_year

    print(survey_period)

    print(output_file_name)

    time.sleep(60)

    # connect to the output file name selection window and enter the name from above and save the file.

    app5 = Application(backend="uia").connect(title_re='Select an output file name', visible_only=True)
    print("\nConnect app5 filename is ", output_file_name)
    app5.window(title_re='Select an output file name').type_keys(output_file_name, with_spaces=True)
    app5.window(title_re='Select an output file name').Save.click()
    del app5

    time.sleep(60)

    # Wait patiently

    while nrcd_running_check("2"):
        logger.info('waiting for ' + local_authority + 'QA output to finish')
        time.sleep(90)

    logger.info(local_authority + ' QA output complete')

    return  # back to main code block


dataloading(app, file_path_variable)

assign_la(app, file_path_variable)

fitting(app)

assign_urb_rural(app)

scanner_qa(app, file_path_variable)

logger.info('End of the run')
