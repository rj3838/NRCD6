from pandas import DataFrame
# from pywinauto.application import Application
import glob
# import logging
# import os
# import sys
# import time
# import tkinter
# from tkinter import filedialog
# import pywinauto
import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename



# Function to find all the files for that authority

def fun_authority_file_search(file_path_variable: str):
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


def file_loading(load_file: str):
    loaded_df = pd.read_csv(load_file, dtype={"SectionLabel": "str", "SectionID": "str"}, low_memory=False)
    return loaded_df


def func_data_calculation(initial_thin: DataFrame,
                          match_thin: DataFrame,
                          match_columns,
                          match_class,
                          initial_total_chainage):
    print(match_columns)
    print(match_class)

    initial_chainage = initial_thin['Chainage'].sum()

    # Create a DF to work with, matching on the columns we are interested in.

    working_df = initial_thin.merge(match_thin,
                                        on=match_columns,
                                        suffixes=('_initial', '_match'),
                                        how='left')

    # select only the rows with the class of road we are looking for
    working_df = working_df.loc[working_df["Class_initial"].isin(match_class)]
    count_of_matching_records = len(working_df)
    print(count_of_matching_records, " rows")

    # calculate the chainage of the initial class then the matching class
    # matching class will be the same oo less

    chainage_of_initial_class = working_df['Chainage_initial'].sum()
    chainage_of_matching_class = working_df['Chainage_match'].sum()

    #
    # percentage_difference_to_total = (abs(initial_total_chainage - chainage_of_matching_class) /
    #                          ((initial_total_chainage + chainage_of_matching_class) / 2)) * 100
    # print(percentage_difference_to_total, 'Percentage difference to year')
    # percentage_difference = (abs(chainage_of_initial_class - chainage_of_matching_class) /
    #                          ((chainage_of_initial_class + chainage_of_matching_class) / 2)) * 100
    # print(chainage_of_initial_class, ' current metres ', chainage_of_matching_class, ' previous metres ',
    #       percentage_difference, '% diff of selection')
    #
    # smallest_chainage = min([chainage_of_initial_class, chainage_of_matching_class])
    # largest_chainage = max([chainage_of_initial_class, chainage_of_matching_class])
    percentage_chainage_match = (chainage_of_matching_class / chainage_of_initial_class) * 100
    # percentage_chainage_match = (chainage_of_matching_class / initial_total_chainage) * 100

    return percentage_chainage_match


def fun_data_comparison(initial_thin: DataFrame,
                        match_thin: DataFrame,
                        match_columns: list,
                        match_class: list) -> float:
    #    print(match_columns, ' ', match_class)

    percentage_difference_to_total = func_data_calculation(initial_df,
                                                           match_thin,
                                                           match_columns,
                                                           match_class,
                                                           initial_total_chainage)

    # return  ((lambda: -9999, percentage_difference_to_total)[percentage_difference_to_total > 100]())

    # return percentage_difference_to_total

    # if percentage_difference_to_total > 100:
    #     return -9999
    # else:
    return percentage_difference_to_total


def create_thin_df(long_df: DataFrame):
    thin_df = long_df.groupby(['SectionLabel', 'SectionID', 'Lane', 'Class'])['Chainage'] \
        .max().reset_index()

    return thin_df


# THIS IS THE MAIN PROC

# select the initial (current) file to match

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
initial_file: str = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
print(initial_file)

files_to_process: list = fun_authority_file_search(initial_file)

initial_df: pd.DataFrame = file_loading(initial_file)

# probably don't need to sort but belt and braces...

sort_order: list = ['SectionLabel', 'SectionID', 'Lane', 'Class', 'UR', 'Chainage']

# group the data we are comparing to into the selection columns and the chainage for the selection
# which will be the max(chainage) for the selection columns

initial_thin = create_thin_df(initial_df)

initial_total_chainage = initial_thin['Chainage'].sum()
print('current total chainage', initial_total_chainage)

# for each file in the list of previous year files

for match_file in files_to_process:

    match_df = file_loading(match_file)

    # match_thin = match_df.groupby(sort_order).max('Chainage')

    print(initial_file)
    print(match_file)

    # reduce the columns we are dealing with so
    # group the data we are comparing to into the selection columns and the chainage for the selection
    # which will be the max(chainage) for the selection columns

    match_thin = create_thin_df(match_df)

    match_total_chainage = match_thin['Chainage'].sum()
    print('previous total chainage', match_total_chainage)

    # match_thin.sort_values(by=sort_order, inplace=True)

    match_all: list = ['SectionLabel', 'SectionID', 'Lane']
    match_label_lane: list = ['SectionLabel', 'Lane']
    match_id_lane: list = ['SectionID', 'Lane']
    match_column_sets = [match_all, match_label_lane, match_id_lane]

    match_all_class: list = ['A', 'B', 'C']
    match_a_class: list = ['A']
    match_b_and_c_class: list = ['B', 'C']
    match_class_sets = [match_all_class, match_a_class, match_b_and_c_class]

    column_list = ['match_all', 'match_label_lane', 'match_id_lane']
    row_list = ['match_all_class', 'match_a_class', 'match_b_and_c_class']
    year_results_df = pd.DataFrame()

    for match_column in match_column_sets:
        column = str(match_column)
        for match_class in match_class_sets:
            row = str(match_class)
            percentage_difference_to_total = fun_data_comparison(initial_thin,
                                                                 match_thin,
                                                                 match_column,
                                                                 match_class)

            # year_results_df.set_value(column, row, percentage_difference_to_total)
            year_results_df.at[column, row] = percentage_difference_to_total

            # row_index = year_results_df[percentage_difference_to_total].index.values[j]
            # value = df.at[row_index, col_index]
            # year_results_df.at(column, row, percentage_difference_to_total)

    print(initial_file)
    print(match_file)
    print(year_results_df.to_string())
