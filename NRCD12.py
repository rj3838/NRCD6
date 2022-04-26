import logging
import os
import sys
import time
import tkinter
from tkinter import filedialog
import pandas as pd
from create_batchfile import batchfile_creation
# from commonDataCheck import commonDataCheck


def fun_directory_selector(request_string: str, selected_directory_list: list, search_directory):
    directory_path_string = filedialog.askdirectory(initialdir=search_directory, title=request_string)

    if len(directory_path_string) > 0:
        selected_directory_list.append(directory_path_string)
        fun_directory_selector('Select the next Local Authority Directory or Cancel to end', selected_directory_list,
                               os.path.dirname(directory_path_string))

    return selected_directory_list


# def nrcd_running_check(app):

def nrcd_running_check(number_of_windows: str):
    from pywinauto import Application

    # Check the taskbar to find the number of running applications. Anything different from the number_of_windows means
    # that the function returns false and the calling process will wait

    tb = Application(backend="uia").connect(title='Taskbar')  # backend is important!!!
    running_apps = tb.Taskbar.child_window(title="Running applications", control_type="ToolBar")

    nrcd_windows = "NRCD - " + number_of_windows + " running windows"
    print('Looking for', nrcd_windows)

    print([w.window_text() for w in running_apps.children()])

    if nrcd_windows in [w.window_text() for w in running_apps.children()]:
        wait_flag = True
    else:
        wait_flag = False

    # try to find out if the window exists. If it has set True (wait) else False (carry on)
    # import pywinauto as pwa

    # if it all goes wrong then wait anyway...
    # wait_flag = True

    # try to connect to the window
    # try:
    #     app_to_wait_for = pwa.Application(backend="uia").connect(title='',
    #                                                              control_type='Window',
    #                                                              visible_only=True)
    #     wait_flag = True
    #     del app_to_wait_for
    #
    # # if it throws an error the window isn't there. Messy solution but  it works.
    # except pwa.ElementNotFoundError as error:
    #     time.sleep(30)
    #     try:
    #         # del app_to_wait_for
    #         app_to_wait_for = pwa.Application(backend="uia").connect(title='',
    #                                                                  control_type='Window',
    #                                                                  visible_only=True)
    #         wait_flag = True
    #         del app_to_wait_for
    #
    #     except pwa.ElementNotFoundError as error:
    #         wait_flag = False

    return wait_flag


def data_loading(app, file_path_variable):
    # from pywinauto.controls.win32_controls import ComboBoxWrapper
    # from pywinauto.controls.uia_controls import ToolbarWrapper
    time.sleep(30)

    # click on the loading button

    app.window(best_match='National Roads Condition Database - Version*', top_level_only=True)\
        .child_window(best_match='Loading').click()

    # if the file_path_variable directory contains a file 'BatchList' use the 'Select Batch File'
    # else use 'Create Batch file'
    filename = file_path_variable + '/BatchList.txt'

    if not os.path.isfile(filename):

        print("\n creating the BatchList.txt")

        # calling the module to create the batch file

        batchfile_creation(filename)

        # else : - the batch file is there so use it.

    print("\nfile name exists using Select", filename)

    time.sleep(30)

    app.window(best_match='National Roads Condition Database - Version', top_level_only=True) \
        .child_window(best_match='Select Batch file').click()

    filename = filename.replace('/', '\\')
    print("\nfile exists", filename)
    time.sleep(15)

    # check existence of the app2 variable if it is there destroy it as connecting to the file selection
    # is going to create it, and it could get messy if it's still there from processing the previous LA.
    # try:
    #    app2
    # except NameError:
    #    print('app2 not used')
    # else:
    #    del app2

    # Connect to the selection window and enter the batch file name.
    app2 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
    print("\nConnect app2 filename is ", filename)

    app2.window(title_re='Select a batch file') \
        .File_name_Edit.set_text(filename)

    time.sleep(15)

    batch_split_button2 = app2.window(title_re='Select a batch file') \
        .child_window(auto_id='1', control_type="SplitButton")
    from pywinauto.controls.win32_controls import ButtonWrapper
    ButtonWrapper(batch_split_button2).click()

    del app2

    # if the file_path_variable directory string contains WDM
    if "WDM" in file_path_variable:
        # then Survey Contractor = WDM
        survey_contractor: str = "WDM"

        # else if the directory string contains 'G-L' select survey contractor 'Ginger-Lehmann'
    elif "G-L" in file_path_variable:
        survey_contractor: str = "Ginger-Lehmann"

        # else if the directory string contains 'PTS' select survey contractor 'Ginger-Lehmann'
    elif "PTS" in file_path_variable:
        survey_contractor: str = "PTS"

        # it's not one of those we know about. Should another contractor start surveying then add another 'elif'
    else:
        survey_contractor = "Unknown contractor"

    print(survey_contractor)

    from pywinauto.controls.win32_controls import ComboBoxWrapper

    surveyor_combobox = app.window(best_match='National Roads Condition Database') \
        .child_window(best_match='Survey Contractor', control_type='Group') \
        .child_window(auto_id="3", control_type="ComboBox")

    ComboBoxWrapper(surveyor_combobox).select(survey_contractor)

    # click 'OK' to close the data loading window as we have all the appropriate details in the window.

    app.window(best_match='National Roads Condition Database') \
        .child_window(best_match='OK').click()

    # back on the NRCD main screen, click the process radio button then the actual 'Process' button

    app.window(best_match='National Roads Condition Database - Version *',
               top_level_only=True).child_window(best_match='Process').click()

    # The log entry contains the first file to be loaded (the rest will not appear and NRCD uses the
    # batch file to find them

    logger.info('Starting loading with ' + filename)

    time.sleep(60)

    # wait for the loading to finish. It checks the number of windows open for NRCD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check('2'):
        logger.info('waiting for loading to complete')
        time.sleep(300)

    logger.info('loading completed')

    return  # back to main


def assign_la(app, file_path_variable):
    logger.info('starting local authority assignment')
    time.sleep(60)

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Attributes').click()
    time.sleep(15)

    group_control = app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Local Authority Attribute", auto_id="3", control_type="Group")

    group_control.child_window(title="Select Batch File", auto_id="7", control_type="Button") \
        .click()

    time.sleep(30)

    filename = file_path_variable + '/BatchList.txt'
    print("\nfilename is ", filename)
    logger.info('File name is ' + filename)
    filename = filename.replace('/', '\\')

    time.sleep(15)

    app4 = Application(backend="uia").connect(title_re='Select a batch file', visible_only=True)
    print("\nConnect app4 filename is ", filename)
    logger.info('Connecting to the batch file selection with ' + filename)

    time.sleep(60)

    # split filename into the directory path and the filename (because Micky soft has changed the window !)
    # directory_path = os.path.dirname(filename)

    app4.window(title_re='Select a batch file') \
        .File_name_Edit.set_text(filename)

    batch_split_button1 = app4.window(title_re='Select a batch file') \
        .child_window(auto_id='1', control_type="SplitButton")
    from pywinauto.controls.win32_controls import ButtonWrapper
    ButtonWrapper(batch_split_button1).click()

    time.sleep(60)

    del app4

    # use the location of the BatchList.txt to extract the local authority. It's two from the end.

    local_authority = filename.split("\\")[-2]

    logger.info('starting to assign attributes for ' + local_authority)

    # Load the human name to database name csv table

    lookup_table = '//trllimited/data/INF_ScannerQA/Audit_Reports/Tools/LA_to_NRCD_Name.csv'

    la_lookup_table = pd.read_csv(lookup_table, index_col='Audit_Report_name', dtype={'INDEX': str})

    la_db_name: str = la_lookup_table.loc[local_authority, "NRCD_Name"]

    logger.info('DB name for the LA is ' + la_db_name)

    time.sleep(15)

    # print(la_db_name)

    batch_combobox1 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='6', control_type="ComboBox").wait('exists enabled visible ready')

    import pywinauto.controls.win32_controls
    pywinauto.controls.win32_controls.ComboBoxWrapper(batch_combobox1).select(la_db_name)

    batch_combobox2 = app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(best_match='Local Authority Attribute', control_type="Group") \
        .child_window(auto_id='4', control_type="ComboBox")  # .wait('exists enabled visible ready')

    pywinauto.controls.win32_controls.ComboBoxWrapper(batch_combobox2).select(" 2021/22")

    time.sleep(15)

    # print(survey_year)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Assign Local Authority", auto_id="5", control_type="Button") \
        .click()

    logger.info('waiting for LA assignment to complete')

    time.sleep(60)

    #  the following contains 10000 seconds. This is to stop the wait timing out, it retries each 90 secs
    #  but the attributes should be finished in under 6 hours... or even one hour.
    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title='NRCD', control_type="Window") \
        .child_window(title="OK", auto_id="2", control_type="Button") \
        .wait("exists ready", timeout=10000, retry_interval=90)

    time.sleep(120)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title='NRCD', control_type="Window") \
        .child_window(title="OK", auto_id="2", control_type="Button").click()

    time.sleep(60)

    app.window(best_match='National Roads Condition Database - Version*') \
        .child_window(title="Exit", auto_id="11", control_type="Button").click()

    return


def fitting(app):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper

    # Turn all the process radio buttons off. Not pretty but it works

    main_screen_process_checkbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_process_checkbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_process_checkbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_process_checkbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    ButtonWrapper(main_screen_process_checkbox).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox2).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox3).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox4).uncheck_by_click()

    del main_screen_process_checkbox
    del main_screen_process_checkbox2
    del main_screen_process_checkbox3
    del main_screen_process_checkbox4

    logger.info('starting fitting')

    time.sleep(30)

    # click fitting to start the next section

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Fitting').click()

    time.sleep(10)

    # set the checkboxes for fitting the data using a grid.

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

    ComboBoxWrapper(batch_combobox24).select(" 2021/22")

    time.sleep(5)

    # with everything set click the OK button to return to the main window.

    main_screen.child_window(title="OK", auto_id="11", control_type="Button").click()

    time.sleep(60)

    # the 'process' click box is already set so click on the main screen Process button at the bottom
    # then wait for the fitting to complete.

    main_screen.Process.click()

    time.sleep(600)

    # wait for the loading to finish. It checks the number of windows open for NRCD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check('2'):
        logger.info('waiting for Fitting to complete')
        time.sleep(300)

    logger.info('fitting complete')

    time.sleep(60)

    # go back to the main code block.
    return


def assign_urb_rural(app):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper

    time.sleep(60)

    main_screen = app.window(best_match='National Roads Condition Database - V*')

    time.sleep(60)

    # on the main screen turn all the checkboxes off.

    main_screen_process_checkbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_process_checkbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_process_checkbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_process_checkbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    time.sleep(60)

    ButtonWrapper(main_screen_process_checkbox).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox2).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox3).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox4).uncheck_by_click()

    del main_screen_process_checkbox
    del main_screen_process_checkbox2
    del main_screen_process_checkbox3
    del main_screen_process_checkbox4

    logger.info('starting urban & rural attributes')

    time.sleep(30)

    # start the attributes section.

    app.window(best_match='National Roads Condition Database',
               top_level_only=True).child_window(best_match='Attributes').click()

    time.sleep(30)

    # to set the urban rural attribute the selections are all made in the 'Attributes Options' part of the window.
    # wait for it to exist.

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Attributes Options", auto_id="12", control_type="Group") \
        .wait("exists ready", timeout=90, retry_interval=60)

    # set group_control to that part of the window.

    group_control = app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Attributes Options", auto_id="12", control_type="Group")

    # group_control.print_control_identifiers()

    # Select SCANNER and the survey year and click on OK

    ComboBoxWrapper(group_control.child_window(auto_id="14",
                                               control_type="ComboBox")) \
        .select("SCANNER")

    ComboBoxWrapper(group_control.child_window(auto_id="15",
                                               control_type="ComboBox"))\
        .select(" 2021/22")

    app.window(best_match='National Roads Condition Database').child_window(best_match='OK').click()

    # back to the main window where the process click box is already set by the NRCD and where it's waiting
    # for the process button at the bottom to be clicked.

    main_screen.Process.click()

    time.sleep(60)

    # after an appropriate time start waiting for the processing to finish.

    # wait for the window to close. It checks the number of windows open for NRCD.exe. If these are less than
    # two the section is complete, otherwise it loops.

    while nrcd_running_check('2'):
        logger.info('waiting for Urban/Rural Attributes to complete')
        time.sleep(300)

    logger.info('U/R attributes complete')

    return  # to the main code block


def scanner_qa(app, file_path_variable):
    from pywinauto.controls.win32_controls import ComboBoxWrapper
    from pywinauto.controls.win32_controls import ButtonWrapper
    import re  # there is a regex search below !

    time.sleep(150)

    main_screen = app.window(best_match='National Roads Condition Database - V*')

    # turn all the process check boxes off.

    main_screen_process_checkbox = main_screen.child_window(title="Process", auto_id="15", control_type="CheckBox")
    main_screen_process_checkbox2 = main_screen.child_window(title="Process", auto_id="16", control_type="CheckBox")
    main_screen_process_checkbox3 = main_screen.child_window(title="Process", auto_id="17", control_type="CheckBox")
    main_screen_process_checkbox4 = main_screen.child_window(title="Process", auto_id="18", control_type="CheckBox")

    time.sleep(60)

    ButtonWrapper(main_screen_process_checkbox).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox2).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox3).uncheck_by_click()
    ButtonWrapper(main_screen_process_checkbox4).uncheck_by_click()

    filename = file_path_variable.replace('/', '\\')
    print(filename)
    # time.sleep(15)

    # use the location of the file_path_variable to extract the local authority, nation and year.

    local_authority = filename.split("\\")[-1]
    nation = filename.split("\\")[-2]
    year: str = re.search('[1-2][0-9]{3}', filename).group(0)  # regex search

    # logger.info('starting Scanner QA output for ' + local_authority)
    # print(local_authority)
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

    time.sleep(30)

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Survey QA Options", auto_id="9", control_type="Group") \
        .wait("exists ready", timeout=90, retry_interval=60)

    group_control = app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="Survey QA Options", auto_id="9", control_type="Group")

    # group_control.print_control_identifiers()

    # exclude the previous year and the U roads (uncheck_by_click) then select the LA abd survey year.

    ButtonWrapper(group_control
                  .child_window(title="Include Previous Year",
                                control_type="CheckBox")).uncheck_by_click()

    ButtonWrapper(group_control
                  .child_window(title="U Roads",
                                control_type="CheckBox")).uncheck_by_click()

    ComboBoxWrapper(group_control.child_window(auto_id="24",
                                               control_type="ComboBox")).select(la_db_name)

    ComboBoxWrapper(group_control.child_window(auto_id="25",
                                               control_type="ComboBox")).select(" 2021/22")

    # Export the data

    group_control.child_window(auto_id="26",
                               title="Export QA Data").click()

    # build output file name.
    # LIVE
    output_file_name = os.path.normpath("//trllimited/data/INF_ScannerQA/Audit_Reports/NRCD Data/" + nation + "/" +
                                        local_authority + "_" + year + ".csv")
    # test
    # output_file_name = os.path.normpath("C:/Users/rjaques/temp/Data/" + nation + "/" + local_authority + "_" +
    #                                    year + ".csv")

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
    time.sleep(30)

    # selected_filename = os.path.basename(output_file_name)
    app5.window(title_re='Select an output file name') \
        .File_name_Edit.set_text(output_file_name)

    batch_split_button1 = app5.window(title_re='Select an output file name') \
        .child_window(title='Save', auto_id='1', control_type="Button")
    from pywinauto.controls.win32_controls import ButtonWrapper
    ButtonWrapper(batch_split_button1).click()

    del app5

    time.sleep(60)

    logger.info('waiting for ' + local_authority + ' QA output to finish')

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="NRCD") \
        .wait("exists ready", timeout=10000, retry_interval=60)

    app.window(best_match='National Roads Condition Database - Version *') \
        .child_window(title="NRCD").OK.click()

    # Wait patiently

    # while nrcd_running_check("2"):
    # logger.info('waiting for ' + local_authority + 'QA output to finish')
    # time.sleep(90)

    logger.info(local_authority + ' QA output complete')

    return  # back to main code block


sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED

root = tkinter.Tk()
root.withdraw()

print('Start')

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
                    filename='//trllimited/data/INF_ScannerQA/AutoNRCD_log.txt',
                    filemode='a+')

logger = logging.getLogger('autoNRCD')
# logger.setLevel(logging.INFO)

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

# root_win = tkinter.Tk()
# root.withdraw()

directories_to_process = list()

directories_to_process: list = fun_directory_selector('Please select a directory containing the data',
                                                      directories_to_process,
                                                      curr_dir)

# noinspection DuplicatedCode
print("\n", directories_to_process)

if len(directories_to_process[0]) > 0:

    file_path_variable = str()

    for file_path_variable in directories_to_process:
        if len(file_path_variable) > 0:
            logger.info('Working with ' + file_path_variable)

# logger.info('Working with ' + file_path_variable)
# file_path_variable = filedialog.askdirectory(initial dir=curr_dir,
# title='Please select a directory containing the data')
# print("\n file_path_variable = ", file_path_variable)

        from pywinauto.application import Application

        # For LIVE
        app = Application(backend="uia").start('C:/Program Files (x86)/NRCD/NRCD.exe')
        # for test
        # app = Application(backend="uia").start('C:/Users/rjaques/Software/NRCD/Current Version/NRCD.exe')

        app.window(best_match='National Roads Condition Database',
                   top_level_only=True).child_window(best_match='SCANNER').click()

        time.sleep(30)
        main_screen = app.window(best_match='National Roads Condition Database - V*')

        # app.window(best_match='', top_level_only=True) \
        # .print_control_identifiers()
        time.sleep(30)

        app.window(best_match='', top_level_only=True) \
            .child_window(best_match='Enter System').click()

        time.sleep(60)

        # call the function handling the buttons on the main NRCD window.

        data_loading(app, file_path_variable)

        assign_la(app, file_path_variable)

        fitting(app)

        assign_urb_rural(app)

        scanner_qa(app, file_path_variable)

        logger.info('End of the run')

        app.window(best_match='National Roads Condition Database - Version *').Exit.click()

        # remove references to the previously opened instance of NRCD

        del app

sys.exit()
