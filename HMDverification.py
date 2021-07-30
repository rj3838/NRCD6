import pandas as pd
import tkinter as Tk
from tkinter import filedialog


# if not file_to_check:
    # print("A", file_to_check)
# Tk.withdraw()  # we don't want a full GUI, so keep the root window from appearing
initial_file: str = filedialog.askopenfilename()  # show an "Open" dialog box and return the path to the selected file
# else:  # use what is passed in
    # print("b", file_to_check)
    # initial_file = file_to_check

# pd.read_csv(initial_file, index_col='Audit_Report_name', dtype={'INDEX': str})
# HMD_to_verify = pd.read_table(initial_file, low_memory=False, lineterminator=";", header="None")
#
# print(HMD_to_verify.head(20))
# print(len(HMD_to_verify))
# print(HMD_to_verify.tail(20))

hdl = open(initial_file)
milist= hdl.read().splitlines()
hdl.close()

HMD_to_verify = pd.DataFrame(milist)

print(HMD_to_verify.head(20))
print(len(HMD_to_verify))
print(HMD_to_verify.tail(20))