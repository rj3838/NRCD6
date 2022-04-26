from pandas import DataFrame
# from pywinauto.application import Application
import glob
import pandas as pd
import re
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename


# Function to find all the files for that authority


def fun_authority_file_search(file_path_variable: str):
    # get a list of the files matching the Authority of the selected file

    file_search_variable: str = file_path_variable.replace('/', '\\')

    # replace the year from the filename for a generic search [1-3][0-9]{3}

    year_str: str = (re.findall(r'20\d{2}', file_search_variable))[-1]
    file_search_variable = file_search_variable.replace(year_str, "*")

    # do a generic search to return a list of files and reverse it
    previous_files_to_match = glob.glob(file_search_variable)

    # remove the first file as it is the initial file (recent year) and reverse it
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


def data_check_main(file_to_check: str) -> None:

    # if the file to check string is empty
    if not file_to_check:
        # print("A", file_to_check)
        Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        initial_file: str = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    else:  # use what is passed in
        # print("b", file_to_check)
        initial_file = file_to_check

    files_to_process: list = fun_authority_file_search(initial_file)

    initial_df: pd.DataFrame = file_loading(initial_file)

    # Extract digit string
    year_str = re.sub(r"\D", "", initial_file)
    #  next year is the last two chars of the year + 1
    year_nxt = str(int(year_str) + 1)
    year_nxt = year_nxt[-2:]
    year_directory = '{}-{}'.format(year_str, year_nxt)
    # print(year_directory)

    directory_name, file_name = os.path.split(initial_file)

    ''' 
            The following should only remove the year ie '-9999' some authorities have more than one word and use
            underscores as the separator so it cannot be split at the underscore
    '''
    # remove _9999.csv from the filename
    authority = re.sub(r'(_\d*\.csv$)', '', file_name)
    print("authority ", authority)
    surveyor: str = ''

    for root, dirs, files in os.walk('//trllimited/data/INF_ScannerQA/2021-22 SCANNER Data'):

        # search the directory to find the authority then extract and return the surveyor
        # print(root)
        if authority in dirs:
            # dirs.remove('CVS')  # don't visit CVS directories
            print(root)
            surveyor = root.split('\\')[-2]
            nation = root.split('\\')[-1]

            if (nation == "Scotland") & (surveyor == "WDM"):
                surveyor = "Scotland_WDM"

            print(surveyor)
            # print(nation)

    # build the output file string

    output_file_string = '//trllimited/data/INF_ScannerQA/Audit_Reports/{}-{}/{}/{}/{}CommonData.txt' \
        .format(year_str,
                year_nxt,
                surveyor,
                authority,
                authority)
    print(output_file_string)

    output_file_directory = '//trllimited/data/INF_ScannerQA/Audit_Reports/{}-{}/{}/{}' \
        .format(year_str,
                year_nxt,
                surveyor,  # nation,
                authority)

    if not os.path.exists(output_file_directory):
        os.makedirs(output_file_directory)

    # group the data we are comparing to into the selection columns and the chainage km for the selection
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

        # column_list = ['match_all', 'match_label_lane', 'match_id_lane']
        # row_list = ['match_all_class', 'match_a_class', 'match_b_and_c_class']
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

        # Print the input file
        # print(initial_file)

        output_str = '//trllimited/data/INF_ScannerQA/Audit_Reports/{}/England/{}'.format(year_directory, file_name)
        print(output_str)

        with open(output_file_string, "a") as output_file:

            # print the output grid to the file
            print("---", file=output_file)
            print(initial_file, file=output_file)
            print(match_file, file=output_file)
            print(year_results_df.to_string(), file=output_file)  # dataframe containing the results


# call the data_check_main

if __name__ == '__main__':
    data_check_main("")
else:
    file_to_check: str
    data_check_main(file_to_check)
