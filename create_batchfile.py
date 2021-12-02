def create_batchfile_main(file_to_check: str) -> None:

    import glob
    import os
    import tkinter
    from tkinter.filedialog import askopenfilename
    from pathlib import PureWindowsPath

    root = tkinter.Tk()
    root.withdraw()

    # if the file to check string is empty
    if not file_to_check:
        tkinter.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        initial_file: str = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    else:  # use what is passed in
        initial_file = file_to_check

    check_directory, check_filename = os.path.split(initial_file)

    if os.path.exists(check_directory):
        check_file = check_directory + "/*.HMD"
        hmd_directory_list = glob.glob(check_file)  # this finds upper and lower case .hmd

        hmd_batch_list_file = check_directory + r'\BatchListTest.txt'
        hmd_batch_list_file = PureWindowsPath(hmd_batch_list_file).__str__()

        with open(hmd_batch_list_file, "w") as output:
            for row in hmd_directory_list:
                s = "".join(map(str, row))
                s = PureWindowsPath(s)
                output.write(str(s) + '\n')

    else:
        print("Directory not found")


if __name__ == '__main__':
    create_batchfile_main('')
else:
    directory_to_check: str
    create_batchfile_main(directory_to_check)
