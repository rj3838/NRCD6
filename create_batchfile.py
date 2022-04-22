def create_batchfile_main(file_to_check: str) -> None:

    import glob
    import os
    import tkinter
    from tkinter.filedialog import askopenfilename
    from pathlib import PureWindowsPath

    root = tkinter.Tk()
    root.withdraw()

    # if the file to check string is empty. If is is ask the user for the directory else use what is passed in.

    if not file_to_check:
        tkinter.Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        initial_file: str = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    else:  # use what is passed in
        initial_file = file_to_check

    check_directory, check_filename = os.path.split(initial_file)

    if os.path.exists(check_directory):
        check_file = check_directory + "/*.HMD"
        hmd_directory_list = glob.glob(check_file)  # this finds upper and lower case .hmd

        # make sure there are hmd files in the directory

        if len(hmd_directory_list) > 0:

            hmd_batch_list_file = check_directory + r'\BatchListTest.txt'
            hmd_batch_list_file = PureWindowsPath(hmd_batch_list_file).__str__()

            with open(hmd_batch_list_file, "w") as output:
                for row in hmd_directory_list:
                    s = "".join(map(str, row))
                    s = PureWindowsPath(s)
                    output.write(str(s) + '\n')
        else:
            print("No HMD files found in ", hmd_directory_list )

    else:
        print(check_directory, " Directory not found")

# Main proc. Check it's running in isolation or has been called.
# directory_to_check will be passed from calling proc
def batchfile_creation(directory_to_check) -> None:

    if __name__ == '__main__':
        create_batchfile_main('')
    else:
        # directory_to_check: str
        create_batchfile_main(directory_to_check)
