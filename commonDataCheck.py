from pandas import DataFrame
# from pywinauto.application import Application
import glob
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

    year_str: str = (re.findall(r'20[0-9]{2}', file_search_variable))[-1]
    file_search_variable = file_search_variable.replace(year_str, "*")

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


def func_data_calculation(initial_thin_fdc: DataFrame,
                          match_thin_fdc: DataFrame,
                          match_columns_fdc: list,
                          match_class_fdc: list) -> float:
    """
    :param match_class_fdc:
    :param match_thin_fdc:
    :param initial_thin_fdc:
    :type match_columns_fdc: object
    """

    initial_class_df = initial_thin_fdc.loc[initial_thin_fdc["Class"].isin(match_class_fdc)]

    match_class_df = match_thin_fdc.loc[match_thin_fdc["Class"].isin(match_class_fdc)]

    # Create a DF to work with, matching on the columns we are interested in.
    match_columns_for_merge: list = match_columns_fdc
    match_columns_for_merge.append("Chainage")

    match_columns_for_merge = list(dict.fromkeys(match_columns_for_merge))

    # Remove duplicates
    match_class_df = match_class_df.drop_duplicates(subset=match_columns_for_merge, keep='first')

    working_df = pd.merge(initial_class_df,
                          match_class_df,
                          on=match_columns_for_merge,
                          suffixes=('_initial', '_match'),
                          how='inner')

    # select only the rows with the class of road we are looking for
    # initial_thin: DataFrame = create_thin_df(initial_df)

    # count of initial then matching
    count_of_initial_records: int = initial_class_df.shape[0]

    count_of_matching_records: int = working_df.shape[0]

    percentage_match: float = (count_of_matching_records / count_of_initial_records) * 100

    del match_class_df, match_columns_fdc, match_columns_for_merge

    return percentage_match


def fun_data_comparison(initial_thin_comp: DataFrame,
                        match_thin_comp: DataFrame,
                        match_columns_comp: list,
                        match_class_comp: list) -> float:
    comp_percentage_difference_to_total = func_data_calculation(initial_thin_comp,
                                                                match_thin_comp,
                                                                match_columns_comp,
                                                                match_class_comp)

    return comp_percentage_difference_to_total


def create_thin_df(long_df: DataFrame) -> DataFrame:

    thin_df = long_df[['SectionLabel', 'SectionID', 'Lane', 'Class', 'Chainage']]
    # round chainage to the nearest 10m and drop the duplicates if any
    thin_df = thin_df.round({'Chainage': -1})
    thin_df.drop_duplicates(['SectionLabel', 'SectionID', 'Lane', 'Class', 'Chainage'], keep='last', inplace=True)

    return thin_df


#
# THIS IS THE MAIN PROC
#
# select the initial (current) file to match

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

# when running standalone ask for the filenames to start with otherwise take that from the
# calling the common date check. (_name_ is not _main_)

initial_file: str = ""

if __name__ == "__main__":
    initial_file: str = askopenfilename()
else:
    # initial_file: str = passed_in_file: str
    pass

# show an "Open" dialog box and return the path to the selected file

files_to_process: list = fun_authority_file_search(initial_file)

initial_df: pd.DataFrame = file_loading(initial_file)

# probably don't need to sort but belt and braces...

sort_order: list = ['SectionLabel', 'SectionID', 'Lane', 'Class', 'UR', 'Chainage']

# group the data we are comparing to into the selection columns and the chainage for the selection
# initial_thin = pd.DataFrame()
# match_thin = pd.DataFrame()

initial_thin: DataFrame = create_thin_df(initial_df)

# for each file in the list of previous year files

for match_file in files_to_process:

    match_df = file_loading(match_file)

    # reduce the columns we are dealing with so

    match_thin: DataFrame = create_thin_df(match_df)

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
            # the data is in initial thin (current year) and match thin (matching year) so pass it to the calculation.
            percentage_difference_to_total: float = func_data_calculation(initial_thin,
                                                                          match_thin,
                                                                          match_column,
                                                                          match_class)

            year_results_df.at[column, row] = percentage_difference_to_total

    # store the results

    with open("//trllimited/data/INF_ScannerQA/Audit_Reports/outputTest.txt", "a") as output_file:
        print(" ", file=output_file)
        print(initial_file, file=output_file)
        print(match_file, file=output_file)
        print(year_results_df.to_string(), file=output_file)  # dataframe containing the results
