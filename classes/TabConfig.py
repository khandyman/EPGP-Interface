from tkinter import filedialog
import tkinter as tk
import ttkbootstrap as ttk

import __main__
from classes.Helper import Helper

class TabConfig(ttk.Frame):
    # This class creates the Config Tab, responsible for:
    # - Retrieving/Changing the path to the user's EQ log file;
    #   this path is stored in 'config' in the same directory
    #   as the application
    # - Changing and saving the current window theme

    def __init__(self, parent, setup):
        super().__init__(parent)

        # ----- class members -----
        self._helper = Helper()
        self._app = parent
        self._setup = setup

        self.CONFIG_PATH = "config"
        self.NO_LOG_ERROR = "Please choose a log file to continue."
        self.STUBBORN_LOG_ERROR = "Sorry, you cannot run EPGP without a log file. Bye!"
        self.BAD_CONFIG_ERROR = "Invalid config file format. Please delete config\nin EPGP-Interface directory and re-launch app."
        self.INVALID_LOG_ERROR = "Please select only EverQuest \'eqlog\' files."
        self.INVALID_MULE_ERROR = "Please select only Zeal \'Inventory.txt\' files."

        self._list_frame = tk.Frame(self)
        self._list_scroll = ttk.Scrollbar(self._list_frame, orient=tk.VERTICAL)
        self._mule_list_box = tk.Listbox(self._list_frame)

        self._log_file = ttk.StringVar()
        self._app_theme = ttk.StringVar()
        self._mule_list = []

        self.configure(width=880, height=595)
        self.place(x=0, y=0)

        self.create_widgets()
        # Read in the log file path at startup
        self.import_config()

    # ----- getters -----
    def get_log_file(self):
        return self._setup.get_log_file()

    def get_officer(self):
        return self._setup.get_officer()

    def get_mule_list_selection(self):
        return self._mule_list_box.get(tk.ANCHOR)

    def get_mule_list_contents(self):
        return self._mule_list_box.get(0, tk.END)

    def get_app_theme(self):
        return self._app_theme.get()

    # ----- setters -----
    def set_log_file(self, config_value):
        self._log_file.set(config_value)
        self._setup.set_log_file(config_value)

    def set_mule_list_box(self, set_type, inventory_file):
        if set_type == 'insert':
            self._mule_list_box.insert(tk.END, inventory_file)
        else:
            self._mule_list_box.delete(tk.ANCHOR)

        self.set_mule_list(self.get_mule_list_contents())

    def set_mule_list(self, mule_list):
        self._mule_list = mule_list
        self._setup.set_mule_list(mule_list)

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

        self._list_scroll.configure(command=self._mule_list_box.yview)
        self._list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._mule_list_box.configure(font=('Calibri', 12), yscrollcommand=self._list_scroll.set)
        self._mule_list_box.pack(fill=tk.BOTH)

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

    def import_config(self):
        self.set_log_file(self._setup.get_log_file())

        for item in self._setup.get_mule_list():
            self.set_mule_list_box('insert', item)

    def change_log(self, initial):
        # This function changes a user's log file; it writes
        # the file path to config and sets the UI label
        # Parameters: self (inherit from TabConfig parent)
        # Return: boolean

        # record original file path in case of error
        old_file = self._setup.get_log_file()
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
        from main import ep_tab

        if initial is False:
            ep_tab.look_for_raids(initial)
            ep_tab.clear_ep()

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

        self.set_mule_list_box('delete', tk.ANCHOR)
        self.set_mule_list(self.get_mule_list_contents())
        __main__.bank_tab.set_bank_mule_combo()
        self.write_config()

    def add_mule(self):
        # This function removes a mule from the list box
        # then updates config with current list of mules
        # Parameters:  self (inherit from TabConfig parent)
        # Return: none

        file_path = filedialog.askopenfilename(title="Select Mule Inventory File",
                                               filetypes=[("Text files", "*Inventory.txt")])

        if self.validate_add_mule(file_path):
            if file_path:
                # if validation passes and file exists, then
                # insert new mule file path into list box,
                # update drop down on bank tab, and write
                # to config
                self.set_mule_list_box('insert', file_path)
                self.set_mule_list(self.get_mule_list_contents())
                __main__.bank_tab.set_bank_mule_combo()
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
