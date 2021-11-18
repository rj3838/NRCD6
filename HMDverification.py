import pandas as pd
# import tkinter as Tk
import re
from tkinter import filedialog


# if not file_to_check:
# print("A", file_to_check)
# Tk.withdraw()  # we don't want a full GUI, so keep the root window from appearing
initial_file: str = filedialog.askopenfilename()  # show an "Open" dialog box and return the path to the selected file
# else:  # use what is passed in
# print("b", file_to_check)
# initial_file = file_to_check

# print(HMD_to_verify.head(20))
# print(len(HMD_to_verify))
# print(HMD_to_verify.tail(20))

hdl = open(initial_file)
mi_list = hdl.read().splitlines()
hdl.close()

raw_hmdif = pd.DataFrame(mi_list)

print(raw_hmdif.head(20))
print(len(raw_hmdif))
print(raw_hmdif.tail(20))


# HMSTART check
# check for only one HMSTART

def line_counting(data_to_process, search_string):
    # Counts the number of lines in a list with the search string in it.

    count = 0

    for line in data_to_process:
        if search_string in line:
            count += 1

    return count


def record_testing(data_to_process, list_of_strings, correct_count):
    # find out if the number of items in a list containing the string is the same as the correct_count

    for field in list_of_strings:
        if line_counting(data_to_process, field) != 1:
            print(field + " is missing")

    return


list_of_records = ['HMSTART', 'TSTART', 'TEND', 'DSTART', 'DEND', 'HMEND']

record_testing(raw_hmdif, list_of_records, 1)

# check the record counts
# using filter() + lambda
# to get string with substring
# res = re.findall(r'\d+',(list(filter(lambda x: 'HMEND' in x, raw_hmdif))[0]))


def check_record_counts(start_string, end_string, file_to_check):
    start_index = [i for i, item in enumerate(file_to_check) if item.startswith(start_string)][0]
    end_index = [i for i, item in enumerate(file_to_check) if item.startswith(end_string)][0]

    length_between_indexes = len(file_to_check[start_index:end_index]) + 1

    res = re.findall(r'\d+', (list(filter(lambda x: end_string in x, raw_hmdif))[0]))

    # When the record count does not equal the value in the HMEND record shout...

    if int(res[0]) != int(length_between_indexes):
        print('The ' + end_string + ' record count is wrong')

    #    print('Record count OK')

    return


check_record_counts('HMSTART', 'HMEND', raw_hmdif)
check_record_counts('TSTART', 'TEND', raw_hmdif)
check_record_counts('DSTART', 'DEND', raw_hmdif)
