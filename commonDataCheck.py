from pandas import DataFrame
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

initial_df: pd.DataFrame = pd.read_csv(initial_file,
                         dtype={"SectionLabel": "str",
                                "SectionID": "str"})

sort_order: list = ['SectionLabel', 'SectionID', 'Lane', 'Class', 'UR', 'Chainage']

initial_df_short = initial_df.groupby(['SectionLabel', 'SectionID', 'Lane', 'Class'])['Chainage']\
    .max().reset_index()

# initial_df_short = initial_df.loc[:, sort_order]

# initial_df_short.sort_values(by=sort_order, inplace=True)

# remove the initial fie from the list to process

# for each file in the list of previous year files

for match_file in files_to_process:

    match_df = pd.read_csv(match_file,
                           dtype={"SectionLabel": "str",
                                  "SectionID": "str"})
    # match_df_short = match_df.groupby(sort_order).max('Chainage')

    # reduce the columns we are dealing with
    match_df_short = match_df.groupby(['SectionLabel', 'SectionID', 'Lane', 'Class'])['Chainage'] \
        .max().reset_index()

    # match_df_short.sort_values(by=sort_order, inplace=True)


    # join the initial and previous data frames on the following fields
    # SectionLabel
    # SectionID
    # Lane (thats direction)
    # Class
    # UR (thats type)
    # Length
    left_key: list = ['SectionLabel', 'SectionID', 'Lane']

    match_all: list = ['SectionLabel', 'SectionID', 'Lane']
    match_label_dir: list = ['SectionLabel', 'Lane']
    match_id_dir: list = ['SectionID', 'Lane']
    match_a_class: list = ['A']
    match_b_and_c_class: list = ['B', 'C']

    # right_key: list = ['SectionLabel', 'SectionID', 'Lane', 'Class', 'UR', 'Chainage']

    all_merged_df = initial_df_short.merge(match_df_short,
                                          on=match_all,
                                          suffixes=('_initial', '_match'),
                                          how='inner')

    # all_merged_df = initial_df_short.join(match_df_short,
     #                                  on=match_all,
      #                                 suffixes=('_initial', '_match'),
      #                                 how='left')

    # print(merged_df.head())

    # all_merged_df_remove_na = all_merged_df.dropna()

    # remove rows where the two road classes do not match

    all_merged_df_with_matching_class = all_merged_df[(all_merged_df['Class_initial'] == all_merged_df['Class_match'])]


    # print(all_merged_df_remove_na.head())

    count_of_all_matching: int = len(all_merged_df_with_matching_class)
    chainage_of_initial_class = all_merged_df_with_matching_class['Chainage_initial'].sum()
    chainage_of_matching_class = all_merged_df_with_matching_class['Chainage_match'].sum()

    percentage_difference = (abs(chainage_of_initial_class - chainage_of_matching_class) /
                             ((chainage_of_initial_class + chainage_of_matching_class) / 2)) * 100
    df_of_class_a: DataFrame = all_merged_df_with_matching_class.loc[all_merged_df_with_matching_class['Class_initial'].isin(['A']) ]
    count_of_all_matching_a_class: int = len(df_of_class_a)
    # count_of_all_matching_a_class: int = len(all_merged_df_remove_na.loc[(all_merged_df_remove_na['Class_initial'].isin(['A'])) ]) # select A roads
    # count_of_all_matching_a_class: int = len(all_merged_df_remove_na['Class_initial'].isin(['A']))  # select A roads
    df_of_class_b_and_c: DataFrame = all_merged_df_with_matching_class.loc[all_merged_df_with_matching_class['Class_initial'].isin(['B', 'C'])]
    count_of_all_matching_b_and_c_class: int = len(df_of_class_b_and_c)
    # count_of_all_matching_b_ana_c_class: int = len(all_merged_df_remove_na.loc[(all_merged_df_remove_na['Class_initial'].isin(['B', 'C']))])
    # count_of_all_matching_b_ana_c_class: int = len(all_merged_df_remove_na['Class_initial'].isin(['B', 'C']))

    print('Count of initial: ', len(initial_df_short))
    print('Count of match: ', len(match_df_short))
    print('Length of initial', all_merged_df_with_matching_class['Chainage_initial'].sum())
    print(chainage_of_initial_class)
    print(chainage_of_matching_class)
    print('Length of matching', all_merged_df_with_matching_class['Chainage_match'].sum())
    print('percentage difference', percentage_difference)
    print('Count of all matching: ', count_of_all_matching)
    # print('Length of initial', initial_df_short['Chainage'].sum(axis=1))

    print('Count of all matching A class: ', count_of_all_matching_a_class)
    print('Count of all matching B&C class: ', count_of_all_matching_b_and_c_class)

    

    # match initial selected file against the previous

    # record match numbers

