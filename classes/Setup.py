from tkinter import filedialog
import tkinter.messagebox
import ttkbootstrap as ttk
import os.path
import sys

from classes.Helper import Helper

class Setup:
    def __init__(self):
        self.CONFIG_PATH = "config"
        self._log_file = ttk.StringVar()
        self._helper = Helper()
        self._mule_list = []

        self.NO_LOG_ERROR = "Please choose a log file to continue."
        self.STUBBORN_LOG_ERROR = "Sorry, you cannot run EPGP without a log file. Bye!"
        self.BAD_CONFIG_ERROR = ("Invalid config file format. Please delete config\n"
                                 "in EPGP-Interface directory and re-launch app.")
        self.INVALID_LOG_ERROR = "Please select only EverQuest \'eqlog\' files."
        self.RAID_FOUND_ERROR = 'No raids found in log file. You will not be able to retrieve EP data.'

        self.import_config(True)

    def get_log_file(self):
        return self._log_file.get()

    def get_mule_list(self):
        return self._mule_list

    def get_officer(self):
        log_file = os.path.basename(self.get_log_file())
        officer_name = log_file[6:len(log_file)-12]

        return officer_name

    def set_log_file(self, config_value):
        self._log_file.set(config_value)

    def set_mule_list(self, mule_list_values):
        self._mule_list = mule_list_values

    def import_config(self, initial):
        # This function is run at start-up to read in
        # the path to the player's EQ log file
        # Parameters: self (inherit from TabConfig parent)
        # Return: player_log (string)

        # If the 'config' file path does not exist display
        # error and run the change_log function

        if not os.path.isfile(self.CONFIG_PATH):
            tkinter.messagebox.showinfo("No Log Found", self.NO_LOG_ERROR)
            self.change_log()

        # Now check for the 'config' file again; if it still
        # doesn't exist (because user didn't choose a file
        # or something else weird is going on, display
        # an error and shut down the app
        if not os.path.isfile(self.CONFIG_PATH):
            tkinter.messagebox.showinfo("No Log Found", self.STUBBORN_LOG_ERROR)
            sys.exit()
        else:
            # If the config file is found, read in the path
            # to the player's log file, set the label widget
            # to it, then return the path as a string
            config_list = self.read_config()
            mule_list = []

            for item in config_list:
                line_slice = item.split(',')

                if len(line_slice) == 2:
                    if line_slice[0] == 'log':
                        self.set_log_file(line_slice[1])
                    else:
                        mule_list.append(line_slice[1])
                else:
                    self._helper.display_error(self.BAD_CONFIG_ERROR)
                    sys.exit()

            self.set_mule_list(mule_list)
            # self.refresh_raids(initial)

    def change_log(self):
        # This function changes a user's log file; it writes
        # the file path to config and sets the UI label
        # Parameters: self (inherit from TabConfig parent)
        # Return: boolean

        # record original file path in case of error
        old_file = self.get_log_file()
        # Pop an open file dialog
        file_path = filedialog.askopenfilename(title="Select Log File", filetypes=[("Text files", "*pq.proj.txt")])

        # if the file path exists
        if file_path:
            # run validation check
            if self.validate_change_log(file_path):
                try:
                    self.set_log_file(file_path)
                    # self.refresh_raids(initial)
                except UnicodeDecodeError as error:
                    # if error, set log file to old file path
                    self.set_log_file(old_file)

                    print(f"An error occurred: {error}")
                    self._helper.display_error(f'An error occurred:\n{error}')

                    # self.refresh_raids(initial)

                    return error

                self.write_config()

            return file_path
        else:
            return False

    def read_config(self):
        # This function reads the user's config file
        # and creates a list of the lines inside
        # Parameters: none
        # Return: config_list (list)

        with open(self.CONFIG_PATH) as file:
            config_list = [line.strip() for line in file]

        return config_list

    def write_config(self):
        # This function writes log and mule file
        # paths to config. this will be done once
        # at initial startup, and whenever a user
        # changes their log file or add/deletes mules
        # Parameters:  self (inherit from TabConfig parent)
        # Return: boolean

        # create a list and set first index to log file
        config_list = [f'log,{self.get_log_file()}']

        # add any mules in list box to list
        for item in self.get_mule_list():
            config_list.append(f'mule,{item}')

        if self.CONFIG_PATH:
            with open(self.CONFIG_PATH, 'w') as file:
                # write each list index to config
                for item in config_list:
                    file.write(f'{item}\n')

            return True
        else:
            return False

    def validate_change_log(self, file_path):
        # This function validates the change_log function
        # by checking for eqlog in the file path
        # Parameters:  file_path (string)
        # Return: boolean

        if 'eqlog' in file_path:
            return True
        else:
            self._helper.display_error(self.INVALID_LOG_ERROR)
            return False
