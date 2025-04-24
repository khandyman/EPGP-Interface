from tkinter import filedialog
import tkinter as tk
import ttkbootstrap as ttk
import os.path
import sys

from classes.Helper import Helper

class TabConfig(ttk.Frame):
    # This class creates the Config Tab, responsible for:
    # - Retrieving/Changing the path to the user's EQ log file;
    #   this path is stored in 'config' in the same directory
    #   as the application
    # - Changing and saving the current window theme

    def __init__(self, parent, ep_tab, bank_tab):
        super().__init__(parent)

        # ----- class members -----
        self._helper = Helper()
        self._app = parent
        self._ep_tab = ep_tab
        self._bank_tab = bank_tab

        self.CONFIG_PATH = "config"
        self.NO_LOG_ERROR = "Please choose a log file to continue."
        self.STUBBORN_LOG_ERROR = "Sorry, you cannot run EPGP without a log file. Bye!"
        self.BAD_CONFIG_ERROR = "Invalid config file format. Please delete config\nin EPGP-Interface directory and re-launch app."
        self.INVALID_LOG_ERROR = "Please select only EverQuest \'eqlog\' files."
        self.INVALID_MULE_ERROR = "Please select only Zeal \'Inventory.txt\' files."

        self._list_frame = tk.Frame(self)
        self._list_scroll = ttk.Scrollbar(self._list_frame, orient=tk.VERTICAL)
        self._mule_list = tk.Listbox(self._list_frame)

        self._log_file = ttk.StringVar()
        self._app_theme = ttk.StringVar()

        self.configure(width=880, height=595)
        self.place(x=0, y=0)

        self.create_widgets()
        # Read in the log file path at startup
        self.import_config(True)

    # ----- getters -----
    def get_log_file(self):
        return self._log_file.get()

    def get_officer(self):
        log_file = os.path.basename(self.get_log_file())
        officer_name = log_file[6:len(log_file)-12]

        return officer_name

    def get_mule_list_selection(self):
        return self._mule_list.get(tk.ANCHOR)

    def get_mule_list_contents(self):
        return self._mule_list.get(0, tk.END)

    def get_app_theme(self):
        return self._app_theme.get()

    # ----- setters -----
    def set_log_file(self, config_value):
        self._log_file.set(config_value)

    def set_mule_list(self, set_type, inventory_file):
        if set_type == 'insert':
            self._mule_list.insert(tk.END, inventory_file)
        else:
            self._mule_list.delete(tk.ANCHOR)

    def set_app_theme(self, config_value):
        self._app_theme.set(config_value)

    # ----- widget setup -----
    def create_widgets(self):
        # This function creates and places all the widgets
        # on the Configure tab
        # Parameters: self (inherit from TabConfig parent)
        # Return: none

        log_label = ttk.Label(self, text='Log File')
        log_label.place(x=20, y=20, width=100, height=30)

        log_entry = ttk.Entry(self, font='Calibri 12', textvariable=self._log_file,
                              foreground="#A9BDBD", state=tk.DISABLED)
        log_entry.place(x=20, y=50, width=650, height=40)

        change_log_button = ttk.Button(self, text='Change Log', style='primary.Outline.TButton',
                                       command=lambda: self.change_log(False))
        change_log_button.place(x=720, y=122, width=140)

        mule_label = ttk.Label(self, text='Mule Outputfiles')
        mule_label.place(x=20, y=174, width=200, height=30)

        self._list_frame.place(x=20, y=204, width=650)

        self._list_scroll.configure(command=self._mule_list.yview)
        self._list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._mule_list.configure(font=('Calibri', 12), yscrollcommand=self._list_scroll.set)
        self._mule_list.pack(fill=tk.BOTH)

        delete_mule_button = ttk.Button(self, text='Delete Mule', style='primary.Outline.TButton',
                                        command=self.delete_mule)
        delete_mule_button.place(x=570, y=476, width=140)

        add_mule_button = ttk.Button(self, text='Add Mule', style='primary.Outline.TButton',
                                     command=self.add_mule)
        add_mule_button.place(x=720, y=476, width=140)

        # theme_label = ttk.Label(self, text='Theme')
        # theme_label.place(x=20, y=175, width=110, height=30)

        # theme_entry = ttk.Combobox(self, font='Calibri 12', textvariable=self._app_theme)
        # theme_entry.bind("<Button-1>", self.update_theme)
        # theme_entry.configure(values=['cosmo', 'flatly', 'journal', 'lumen', 'minty', 'pulse', 'morph', 'solar',
        #                               'superhero', 'darkly', 'cyborg', 'vapor'])
        # theme_entry.place(x=20, y=205, width=175, height=33)

        # change_theme_button = ttk.Button(self, text='Change Theme', style='primary.Outline.TButton',
        #                                  command=self.update_theme)
        # change_theme_button.place(x=630, y=277, width=120)

    # ----- class functions -----
    # todo: complete theme change functionality
    # def update_theme(self):
    # test = ttk.StringVar(value=app.style.theme_use())
    # test.set()
    # app.style.theme_use('vapor')
    # print(self.get_app_theme())

    def import_config(self, initial):
        # This function is run at start-up to read in
        # the path to the player's EQ log file
        # Parameters: self (inherit from TabConfig parent)
        # Return: player_log (string)

        # If the 'config' file path does not exist display
        # error and run the change_log function

        if not os.path.isfile(self.CONFIG_PATH):
            self.hide_window()
            self._helper.display_error(self.NO_LOG_ERROR)
            self.change_log(initial)

        # Now check for the 'config' file again; if it still
        # doesn't exist (because user didn't choose a file
        # or something else weird is going on, display
        # an error and shut down the app
        if not os.path.isfile(self.CONFIG_PATH):
            self._helper.display_error(self.STUBBORN_LOG_ERROR)
            sys.exit()
        else:
            # If the config file is found, read in the path
            # to the player's log file, set the label widget
            # to it, then return the path as a string
            config_list = self.read_config()

            for item in config_list:
                line_slice = item.split(',')

                if len(line_slice) == 2:
                    if line_slice[0] == 'log':
                        self.set_log_file(line_slice[1])
                    else:
                        self.set_mule_list('insert', line_slice[1])
                else:
                    self.hide_window()
                    self._helper.display_error(self.BAD_CONFIG_ERROR)
                    sys.exit()

            self.refresh_raids(initial)

    def hide_window(self):
        self._app.overrideredirect(True)
        self._app.attributes('-alpha', 0)
        self._app.deiconify()

    def read_config(self):
        # This function reads the user's config file
        # and creates a list of the lines inside
        # Parameters: none
        # Return: config_list (list)

        with open(self.CONFIG_PATH) as file:
            config_list = [line.strip() for line in file]

        return config_list

    def change_log(self, initial):
        # This function changes a user's log file; it writes
        # the file path to config and sets the UI label
        # Parameters: self (inherit from TabConfig parent)
        # Return: boolean

        # record original file path in case of error
        old_file = self.get_log_file()
        # Pop an open file dialog
        file_path = filedialog.askopenfilename(title="Select Log File", filetypes=[("Text files", "*.txt")])

        # if the file path exists
        if file_path:
            # run validation check
            if self.validate_change_log(file_path):
                try:
                    self.set_log_file(file_path)
                    self.refresh_raids(initial)
                except UnicodeDecodeError as error:
                    # if error, set log file to old file path
                    self.set_log_file(old_file)

                    print(f"An error occurred: {error}")
                    self._helper.display_error(f'An error occurred:\n{error}')

                    self.refresh_raids(initial)

                    return error

                self.write_config()

            return True
        else:
            return False

    def refresh_raids(self, initial):
        # This function checks for raids in user's log
        # file and updates the ep tab drop down
        # Parameters: none
        # Return: none

        if initial is False:
            self._ep_tab.look_for_raids(initial)
            self._ep_tab.clear_ep()

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
        for item in self.get_mule_list_contents():
            config_list.append(f'mule,{item}')

        if self.CONFIG_PATH:
            with open(self.CONFIG_PATH, 'w') as file:
                # write each list index to config
                for item in config_list:
                    file.write(f'{item}\n')

            return True
        else:
            return False

    def delete_mule(self):
        # This function removes a mule from the list box
        # then updates config with current list of mules
        # Parameters:  self (inherit from TabConfig parent)
        # Return: none

        self.set_mule_list('delete', tk.ANCHOR)
        self._bank_tab.set_bank_mule_combo()
        self.write_config()

    def add_mule(self):
        # This function removes a mule from the list box
        # then updates config with current list of mules
        # Parameters:  self (inherit from TabConfig parent)
        # Return: none

        file_path = filedialog.askopenfilename(title="Select Mule Inventory File",
                                               filetypes=[("Text files", "*.txt")])

        if self.validate_add_mule(file_path):
            if file_path:
                # if validation passes and file exists, then
                # insert new mule file path into list box,
                # update drop down on bank tab, and write
                # to config
                self.set_mule_list('insert', file_path)
                self._bank_tab.set_bank_mule_combo()
                self.write_config()

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

    def validate_add_mule(self, file_path):
        # This function validates the add_mule function
        # by checking for Inventory.txt in the file path
        # Parameters:  file_path (string)
        # Return: boolean

        if 'Inventory.txt' in file_path:
            return True
        else:
            self._helper.display_error(self.INVALID_MULE_ERROR)
            return False