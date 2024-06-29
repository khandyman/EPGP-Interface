import os.path
import platform
import time
import tkinter.messagebox

import requests
import re
import sys
import tkinter as tk
from datetime import datetime
from tkinter import *
from tkinter import filedialog
import tksheet
import ttkbootstrap as ttk
import ttkbootstrap.dialogs
from file_read_backwards import FileReadBackwards
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# This program was created for use by the leadership of the Seekers of Souls (SoS) guild on the
# Project Quarm emulated EverQuest server.  It is a front-end interface for the back end EPGP Log
# that SoS uses to store attendance records and points earned as well as distribute loot obtained
# during raids and guild events.  The EPGP-Interface was designed specifically to automate as much
# of the process of taking attendance, determining loot winners, and entering loot awarded as possible.
# Where the process could not be automated an attempt was made to streamline it and make it easier.


# ---------- constants ----------
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = "1DJm8BhaaxZlwLea0ZVZcyVUhMvQ8_VcmIyahBijajjg"  # official EPGP log

# A1 notation ranges for Google Sheets API interactions
GP_LOG_RANGE = "GP Log!C2:C"
EP_LOG_RANGE = "EP Log!B3:B"
PRIORITY_TYPE_RANGE = "EP Log!R3:R"
CHARACTER_GEAR_RANGE = "Get Priority!B3:C"
LOOT_WINNER_RANGE = "Get Priority!G4:I4"
GET_PRIORITY_RANGE = "Get Priority!E3:E"
COUNT_PRIORITY_RANGE = "Get Priority!D3:D"
BID_LEVELS_RANGE = "Get Priority!H19:H"
RAIDER_LIST_RANGE = "Totals!C4:C"
EFFORT_POINTS_RANGE = "Point Values!B4:B"
GEAR_POINTS_RANGE = "Point Values!H4:H"
LOOT_NAMES_RANGE = "GP Log!E2:E"
LEVEL_RANGE = "Totals!E4:E"

# hard-coded lists
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
CLASSES = ['Bard', 'Beastlord', 'Cleric', 'Druid', 'Enchanter', 'Magician', 'Monk', 'Necromancer',
           'Paladin', 'Ranger', 'Rogue', 'Shadowknight', 'Shaman', 'Warrior', 'Wizard']
RACES = ['Barbarian', 'Dark Elf', 'Dwarf', 'Erudite', 'Gnome', 'Half Elf', 'Halfling', 'High Elf',
         'Human', 'Iksar', 'Ogre', 'Troll', 'Vah Shir', 'Wood Elf']

# validation messages
SCAN_ERROR = 'Please select an available time stamp'
EMPTY_EP_ERROR = 'Please enter at least one full line of data in the EP sheet and try again.'
SPELLING_EP_ERROR = ' ] is misspelled or may be an alt. Please correct and try again'
TYPE_EP_ERROR = 'Please select a point type and try again.'
ENTER_GP_ERROR = 'Please enter a date, character, loot and gear level.'
FIND_WINNER_ERROR = 'Please open bidding and enter at least two characters'
COPY_WINNER_ERROR = 'Please find a winner first'
READ_WINNER_ERROR = 'Problem reading winner from Get Priority tab. Did you enter any alts by mistake?'
RAID_FOUND_ERROR = 'No raids found in log file. You will not be able to retrieve EP data.'
NO_LOG_ERROR = "Please choose a log file to continue."
STUBBORN_LOG_ERROR = "Sorry, you cannot run EPGP without a log file. Bye!"
NO_VOICE_ERROR = "No users found in Guild-Meeting channel"


# ---------- window classes ----------
class Notebook(ttk.Window):
    # root window class; parent of all tab classes

    def __init__(self):
        # ----- window setup -----
        super().__init__(themename='solar')
        self.withdraw()
        self.title('EPGP Interface')
        self.geometry('880x630')
        self.resizable(False, False)
        # main window icon
        self.iconbitmap('WindowsIcon.ico')
        # message box icon
        self.iconbitmap(default='WindowsIcon.ico')

        # Unused - print all colors in theme
        # Colors = self.style.colors
        # for color_label in Colors.label_iter():
        #     color = Colors.get(color_label)
        #     print(color_label, color)

        # todo: revisit ttk style to set style for all widgets at once
        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Calibri', 12))
        style.configure('TCalendar', font=('Calibri', 12))
        style.configure('TLabel', font=('Calibri', 14))
        style.configure('TEntry', font=('Calibri', 12))
        style.configure('primary.Outline.TButton', font=("Calibri", 12))

        # ----- tab setup -----
        # main notebook widget, parent of all frame (tab) widgets
        self.epgp_notebook = ttk.Notebook(self)

    def add_tab(self, class_tab, title):
        # This function creates tabs for the main notebook
        # Parameters: self (inherit from Notebook parent)
        #             class_tab - the name of a class instance
        #             title - the text to display on the window tab
        # Return: none

        self.epgp_notebook.add(class_tab, text=title)
        self.epgp_notebook.place(x=0, y=0, width=880, height=630)

    def run(self):
        # This function executes the window main loop

        self.mainloop()


class Load:
    def __init__(self, parent):
        # This class creates the load screen window.

        self.loading_window = tk.Toplevel(parent)
        self.loading_window.title("Loading...")
        # Hide the Windows title bar
        self.loading_window.overrideredirect(True)

        parent_x_pos = parent.winfo_rootx()
        parent_y_pos = parent.winfo_rooty()
        geometry = '350x150+%d+%d' % (parent_x_pos + 190, parent_y_pos + 180)
        self.loading_window.geometry(geometry)

        self.label_frame = ttk.Frame(self.loading_window)
        self.label_frame.place(x=0, y=30, width=350, height=35)

        self.progress_label = ttk.Label(self.label_frame, text='Loading...', font='Calibri 14')
        self.progress_label.pack()

        self.progress_frame = ttk.Frame(self.loading_window)
        self.progress_frame.place(x=30, y=100, width=285, height=35)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame,
                                            length=250,
                                            variable=self.progress_var,
                                            maximum=100,
                                            style='TProgressbar')
        self.progress_bar.pack(fill='both', expand=True)

    def update_label(self, value):
        # This function updates the text in the label
        self.progress_label.configure(text='')
        self.progress_label.configure(text=value)
        self.loading_window.update_idletasks()

    def update_progress(self, value):
        # This function updates the value of the progress bar
        new_value = self.progress_var.get() + value
        self.progress_var.set(new_value)
        self.loading_window.update_idletasks()


class AddRecord(tk.Toplevel):
    # This class creates the toplevel entry form
    # to add records manually to the ep sheet

    def __init__(self, parent, data_list):
        super().__init__(parent)

        # Obtain parent window position and place
        # this form relative to that
        parent_x_pos = parent.winfo_rootx()
        parent_y_pos = parent.winfo_rooty()
        geometry = '480x240+%d+%d' % (parent_x_pos + 150, parent_y_pos + 180)

        #
        self.transient(parent)
        self.grab_set()
        self.geometry(geometry)
        self.resizable(False, False)

        # Class data members
        self._entry_frame = ttk.Frame(self)
        self._add_month = tk.StringVar()
        self._add_date = tk.StringVar()
        self._add_time = tk.StringVar()
        self._add_year = tk.StringVar()
        self._add_level = tk.StringVar()
        self._add_class = tk.StringVar()
        self._add_name = tk.StringVar()
        self._add_race = tk.StringVar()
        # This will either be an empty list or it will
        # contain data the user is trying to edit
        self._data_list = data_list

        self.populate_fields(data_list)
        self.create_widgets()

    # Public class member getters
    def get_add_month(self):
        return self._add_month.get()

    def get_add_date(self):
        return self._add_date.get()

    def get_add_time(self):
        return self._add_time.get()

    def get_add_year(self):
        return self._add_year.get()

    def get_add_level(self):
        return self._add_level.get()

    def get_add_class(self):
        return self._add_class.get()

    def get_add_name(self):
        return self._add_name.get()

    def get_add_race(self):
        return self._add_race.get()

    # Public class member setters
    def set_add_month(self, value):
        self._add_month.set(value)

    def set_add_date(self, value):
        self._add_date.set(value)

    def set_add_time(self, value):
        self._add_time.set(value)

    def set_add_year(self, value):
        self._add_year.set(value)

    def set_add_level(self, value):
        self._add_level.set(value)

    def set_add_class(self, value):
        self._add_class.set(value)

    def set_add_name(self, value):
        self._add_name.set(value)

    def set_add_race(self, value):
        self._add_race.set(value)

    def populate_fields(self, data_list):
        # This function auto populates some of the fields
        # on the AddRecord form, to save time
        # Parameters: self (inherit from AddRecord parent)
        # Return: none

        # If data_list contains anything, this is an edit
        # existing record rather than an add new record,
        # so pull in the relevant data fields
        if len(data_list) > 0:
            self.set_add_month(data_list[0])
            self.set_add_date(data_list[1])
            self.set_add_time(data_list[2])
            self.set_add_year(data_list[3])
            self.set_add_level(data_list[4])
            self.set_add_class(data_list[5])
            self.set_add_name(data_list[6])
            self.set_add_race(data_list[7])
        # Otherwise, run through some default field
        # population
        else:
            time = datetime.now()

            # If already data in ep sheet, match month, date,
            # time, and year to existing data
            if len(ep_tab.get_ep_sheet()) > 0:
                self.set_add_month(ep_tab.get_ep_sheet_cell("A1"))
                self.set_add_date(ep_tab.get_ep_sheet_cell("B1"))
                self.set_add_time(ep_tab.get_ep_sheet_cell("C1"))
                self.set_add_year(ep_tab.get_ep_sheet_cell("D1"))
            else:
                # Use date/time formatting to make sure data
                # matches EPGP Log
                self.set_add_month(time.strftime('%b'))
                self.set_add_date(time.strftime('%d'))
                self.set_add_time(time.strftime('%H:%M:%S'))
                self.set_add_year(time.strftime('%Y'))

            # Just set a default level of 50; user will have
            # to modify if needed
            self.set_add_level(sheets.get_max_level())

    def create_widgets(self):
        # This function lays out the widgets on the form
        # Parameters: self (inherit from AddRecord parent)
        # Return: none

        self._entry_frame.pack(expand=True, fill='both')

        add_month_label = ttk.Label(self._entry_frame, text='Month', font='Calibri 14')
        add_month_label.place(x=10, y=10, width=70)

        add_month_entry = AutocompleteCombobox(self._entry_frame, MONTHS)
        add_month_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_month)
        add_month_entry.place(x=10, y=45, width=70)

        add_date_label = ttk.Label(self._entry_frame, text='Date', font='Calibri 14')
        add_date_label.place(x=90, y=10, width=50)

        dates = [str(x).zfill(1) for x in range(32)]
        add_date_entry = ttk.Combobox(self._entry_frame, values=dates)
        add_date_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_date)
        add_date_entry.place(x=90, y=45, width=50)

        add_time_label = ttk.Label(self._entry_frame, text='Time', font='Calibri 14')
        add_time_label.place(x=150, y=10, width=100)

        add_time_entry = ttk.Entry(self._entry_frame)
        add_time_entry.configure(font='Calibri 12', textvariable=self._add_time)
        add_time_entry.place(x=150, y=45, width=100)

        add_year_label = ttk.Label(self._entry_frame, text='Year', font='Calibri 14')
        add_year_label.place(x=260, y=10, width=60)

        add_year_entry = ttk.Entry(self._entry_frame)
        add_year_entry.configure(font='Calibri 12', textvariable=self._add_year)
        add_year_entry.place(x=260, y=45, width=60)

        add_level_label = ttk.Label(self._entry_frame, text='Level', font='Calibri 14')
        add_level_label.place(x=330, y=10, width=50)

        levels = [str(x).zfill(1) for x in range(35, 66)]
        add_level_entry = ttk.Combobox(self._entry_frame, values=levels)
        add_level_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_level)
        add_level_entry.place(x=330, y=45, width=50)

        add_class_label = ttk.Label(self._entry_frame, text='Class', font='Calibri 14')
        add_class_label.place(x=10, y=90, width=150)

        add_class_entry = AutocompleteCombobox(self._entry_frame, CLASSES)
        add_class_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_class)
        add_class_entry.place(x=10, y=125, width=150)

        add_name_label = ttk.Label(self._entry_frame, text='Name', font='Calibri 14')
        add_name_label.place(x=170, y=90, width=160)

        add_name_entry = AutocompleteCombobox(self._entry_frame, sheets.get_player_list())
        add_name_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_name)
        add_name_entry.bind("<<ComboboxSelected>>", self.get_player_data)
        add_name_entry.bind("<FocusOut>", self.get_player_data)
        add_name_entry.place(x=170, y=125, width=160)
        add_name_entry.focus_set()

        add_race_label = ttk.Label(self._entry_frame, text='Race', font='Calibri 14')
        add_race_label.place(x=340, y=90, width=130)

        add_race_entry = AutocompleteCombobox(self._entry_frame, RACES)
        add_race_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_race)
        add_race_entry.place(x=340, y=125, width=130)

        clear_add_button = ttk.Button(self._entry_frame, text='Clear', style='primary.Outline.TButton',
                                      command=self.clear_form)
        clear_add_button.place(x=175, y=190, width=140)

        submit_button = ttk.Button(self._entry_frame, text='Submit', style='primary.Outline.TButton',
                                   command=self.submit)
        submit_button.place(x=325, y=190, width=140)

    def get_player_data(self, event):
        # This function auto-populates the class and
        # races boxes, according to the name entered
        # in the name box
        # Parameters: self (inherit from AddRecord parent),
        #             event (object generated by event)
        # Return: none
        if self.validate_name():
            self.set_add_race(sheets.get_player_race(self.get_add_name()))
            self.set_add_class(sheets.get_player_class(self.get_add_name()))

    def clear_form(self):
        # This function clears out all the
        # fields on the entry form
        # Parameters: self (inherit from AddForm parent)
        # Return: none

        self.set_add_month('')
        self.set_add_date('')
        self.set_add_time('')
        self.set_add_year('')
        self.set_add_level('')
        self.set_add_class('')
        self.set_add_name('')
        self.set_add_race('')

    def submit(self):
        # This function captures the form data
        # into a list and inserts it into the ep sheet
        # Parameters: self (inherit from AddForm parent)
        # Return: none

        if self.validate_submit():
            # Create the list of form fields
            entry_list = [self.get_add_month(), self.get_add_date(), self.get_add_time(), self.get_add_year(),
                          self.get_add_level(), self.get_add_class(), self.get_add_name()]

            # Special logic to separate elf races into
            # two words to match the EPGP Log
            race = self.get_add_race()

            if race == 'High Elf' or race == 'Dark Elf' or race == 'Wood Elf':
                race = race.split()
                entry_list.append(race[0])
                entry_list.append(race[1])
            else:
                entry_list.append(race)
                entry_list.append('')

            # If data_list is not empty, this is an edit, so get
            # row number from data_list
            if len(self._data_list) > 0:
                index = self._data_list[8]
            else:
                # Otherwise, obtain the last row of the ep sheet
                # so we know where to insert the new data
                last_row = ep_tab.get_ep_rows()
                # Create the index
                index = f'A{last_row + 1}'

                # Add a row to the ep sheet and insert
                # the form items
                ep_tab.add_ep_row()

            ep_tab.add_ep_item(index, entry_list)

            # Close the toplevel form, returning
            # control to the main window
            self.destroy()
            self.update()

    def validate_name(self):
        # This function is a special validation case
        # to make sure the name in the add_name box
        # is correct, in case the user prematurely
        # leaves the combobox and triggers the FocusOut
        # event
        for item in sheets.get_player_list():
            if self.get_add_name() == item:
                return True

        return False

    def validate_submit(self):
        err_type = ''
        validation_flags = [False, False, False, False, False, False, False, False]

        for item in MONTHS:
            if self.get_add_month() == item:
                validation_flags[0] = True
                break

        num_list = [str(x).zfill(1) for x in range(32)]

        for item in num_list:
            if self.get_add_date() == item:
                validation_flags[1] = True
                break

        time_correct = re.search("^\\d{2}:\\d{2}:\\d{2}", self.get_add_time())

        if time_correct:
            validation_flags[2] = True

        year_correct = re.search("^\\d{4}", self.get_add_year())

        if year_correct:
            validation_flags[3] = True

        num_list = [str(x).zfill(1) for x in range(35, 66)]

        for item in num_list:
            if self.get_add_level() == item:
                validation_flags[4] = True
                break

        for item in CLASSES:
            if self.get_add_class() == item:
                validation_flags[5] = True
                break

        for item in sheets.get_player_list():
            if self.get_add_name() == item:
                validation_flags[6] = True
                break

        for item in RACES:
            if self.get_add_race() == item:
                validation_flags[7] = True
                break

        if not validation_flags[0]:
            err_type = 'month'
        elif not validation_flags[1]:
            err_type = 'date'
        elif not validation_flags[2]:
            err_type = 'time'
        elif not validation_flags[3]:
            err_type = 'year'
        elif not validation_flags[4]:
            err_type = 'level'
        elif not validation_flags[5]:
            err_type = 'class'
        elif not validation_flags[6]:
            err_type = 'name'
        elif not validation_flags[7]:
            err_type = 'race'

        err_msg = f'Please enter a valid [ {err_type} ] before submitting'

        for item in validation_flags:
            if item is False:
                display_error(err_msg)
                return False

        return True


class TabEP(ttk.Frame):
    # This class creates the EP Tab, responsible for:
    # - Retrieving log file time stamps and displaying attendance data
    # - Sending attendance data to the EPGP Log
    # - Manual input is also available

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # Notebook.__init__(app)

        # ----- private class members -----

        # Two variables for the time entry widget,
        # one is a tkinter StringVar to access the
        # widget text contents, the other is a
        # direct reference to the widget itself,
        # so that the configure parameter can be
        # accessed and the combobox values changed
        self._ep_time = tk.StringVar()
        self._ep_time_combo = ttk.Combobox(self)
        self._ep_type = tk.StringVar()
        self._ep_sheet = tksheet.Sheet(self)

        self.configure(width=880, height=630)
        self.place(x=0, y=0)

        # build and configure the widget layout
        self.create_widgets()
        self.set_ep_grid()

    # ----- public class member getters -----
    def get_ep_time(self):
        return self._ep_time.get()

    def get_ep_combo_values(self):
        return self._ep_time_combo['values']

    def get_ep_type(self):
        return self._ep_type.get()

    def get_ep_sheet(self):
        return self._ep_sheet.get_sheet_data()

    def get_ep_sheet_cell(self, cell):
        return self._ep_sheet.get_data(cell)

    # ----- public class member setters -----
    def set_ep_time(self, ep_value):
        self._ep_time.set(ep_value)

    def set_ep_combo_values(self, ep_value):
        self._ep_time_combo.configure(values=ep_value)

    def set_ep_type(self, ep_value):
        self._ep_type.set(ep_value)

    def set_ep_sheet(self, ep_value):
        self._ep_sheet.set_sheet_data(ep_value)

    # ----- widget setup -----
    def create_widgets(self):
        # This function configures and places the tkinter
        # widgets onto the EP Tab frame
        # Parameters:  self (inherit from TabEP parent)
        # Return: none

        ep_title = ttk.Label(self, text='EP Entry', font='Calibri 24 bold')
        ep_title.place(x=380, y=0)

        ep_time_label = ttk.Label(self, text='Import Criteria')
        ep_time_label.place(x=110, y=50, width=360, height=30)

        ep_type_label = ttk.Label(self, text='Point Type')
        ep_type_label.place(x=590, y=50, width=170, height=30)

        self._ep_time_combo.configure(font=('Calibri', 12), height=40, textvariable=self._ep_time, values=[])
        self._ep_time_combo.place(x=110, y=80, width=360)

        ep_type_entry = AutocompleteCombobox(self, sheets.get_effort_points())
        ep_type_entry.configure(font=('Calibri', 12), height=40, textvariable=self._ep_type)
        ep_type_entry.place(x=590, y=80, width=170)

        ep_header = ('Month', 'Date', 'Time', 'Year', 'Level', 'Class', 'Name', 'Race', '.')
        (self._ep_sheet.set_header_data(ep_header)
         .set_sheet_data([]))
        # ttkbootstrap to tksheet mappings: table_bg = inputbg, table_fg = inputfg
        (self._ep_sheet.set_options(table_bg='#073642',
                                    table_fg='#A9BDBD',
                                    header_fg='#ffffff',
                                    index_fg='#ffffff')
         .font(('Calibri', 12, 'normal')))
        (self._ep_sheet.set_options(defalt_row_height=30)
         .height_and_width(width=840, height=375))
        self._ep_sheet.enable_bindings(("single_select",
                                        "row_select",
                                        "column_width_resize",
                                        "arrowkeys",
                                        "right_click_popup_menu",
                                        "rc_select",
                                        "rc_insert_row",
                                        "rc_delete_row",
                                        "copy",
                                        "cut",
                                        "paste",
                                        "delete",
                                        "undo",
                                        "edit_cell"))
        # Set a double click to activate the edit record function
        self._ep_sheet.bind('<Double-Button-1>', self.edit_record)
        self._ep_sheet.place(x=20, y=135)

        clear_ep_button = ttk.Button(self, text='Clear', style='primary.Outline.TButton',
                                     command=self.clear_ep)
        clear_ep_button.place(x=120, y=535, width=140)

        add_record_button = ttk.Button(self, text='Add', style='primary.Outline.TButton',
                                       command=self.add_record)
        add_record_button.place(x=270, y=535, width=140)

        refresh_raids_button = ttk.Button(self, text='Refresh', style='primary.Outline.TButton',
                                          command=lambda: self.look_for_raids(False))
        refresh_raids_button.place(x=420, y=535, width=140)

        scan_log_button = ttk.Button(self, text='Import', style='primary.Outline.TButton',
                                     command=self.import_ep_data)
        scan_log_button.place(x=570, y=535, width=140)

        save_ep_button = ttk.Button(self, text='Save', style='primary.Outline.TButton',
                                    command=self.insert_ep)
        save_ep_button.place(x=720, y=535, width=140)

    def edit_record(self, event):
        # This function grabs data off the sheet in the row
        # clicked by the user, and then opens the add record
        # form so the data can be edited
        # Parameters: self (inherit from TabEP parent
        #             event
        # Return: none

        # Only execute if the user double-clicked a row header
        curr_col = self._ep_sheet.get_currently_selected()[1]

        if curr_col == 0:
            curr_row = self._ep_sheet.get_currently_selected()[0]
            curr_month = str(self._ep_sheet.get_cell_data(curr_row, 0))
            curr_date = str(self._ep_sheet.get_cell_data(curr_row, 1))
            curr_time = str(self._ep_sheet.get_cell_data(curr_row, 2))
            curr_year = str(self._ep_sheet.get_cell_data(curr_row, 3))
            curr_level = str(self._ep_sheet.get_cell_data(curr_row, 4))
            curr_class = str(self._ep_sheet.get_cell_data(curr_row, 5))
            curr_name = str(self._ep_sheet.get_cell_data(curr_row, 6))
            curr_race = str(self._ep_sheet.get_cell_data(curr_row, 7))

            # Format the two word races into a single string
            if curr_race == 'High' or curr_race == 'Dark' or curr_race == 'Wood':
                curr_race = curr_race + ' ' + 'Elf'

            data_list = [curr_month, curr_date, curr_time, curr_year, curr_level,
                         curr_class, curr_name, curr_race, curr_row]

            AddRecord(self, data_list)

    # ----- class functions -----
    def set_ep_grid(self):
        # This function sets the column widths
        # and row heights for the main sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        (self._ep_sheet
         .column_width(column=0, width=60)  # month
         .column_width(column=1, width=50)  # date
         .column_width(column=2, width=90)  # time
         .column_width(column=3, width=60)  # year
         .column_width(column=4, width=60)  # level
         .column_width(column=5, width=150) # class
         .column_width(column=6, width=150) # name
         .column_width(column=7, width=100)  # race1
         .column_width(column=8, width=40)  # race2
         .set_options(default_row_height=30))

    def get_ep_rows(self):
        # This function obtains the current
        # number of rows in the ep sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: get_total_rows (int)

        return self._ep_sheet.get_total_rows()

    def add_ep_row(self):
        # This function adds a row to the
        # bottom of the ep sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self._ep_sheet.insert_row()

    def add_ep_item(self, index, values):
        # This function sets the index row equal
        # to the list of values passed in
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self._ep_sheet.set_data(index, data=[values])

    def clear_ep(self):
        # This function clears the text entries
        # and the main sheet, either after a Log
        # insertion or the clear button function
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self.set_ep_time('')
        self.set_ep_type('')

        num_rows = self._ep_sheet.get_total_rows()

        # Run a loop, from max rows down to zero,
        # deleting a row each time
        for i in range(num_rows - 1, -1, -1):
            self._ep_sheet.delete_row(i)

        self.set_ep_grid()

    def insert_ep(self):
        # This function appends ep sheet data, pulled
        # from EQ log attendance taken, to the bottom
        # of the EP Log tab in the EPGP Log
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        # Check for successful validation
        if self.validate_ep_entry():
            # The ep sheet returns a 2d list
            guild_list = self.get_ep_sheet()
            # Need a second list with one element for each
            # element in guild_list; this second list
            # is just the point type values for inserting
            # into column R on the EP Log tab
            point_list = []
            point_type = ep_tab.get_ep_type()

            # Loop through each line in the guild list,
            # appending the point type each time
            for line in guild_list:
                point_list.append([point_type])

            # Obtain the insertion point for the EP Log;
            # i.e., the first empty cell in column B
            row_count = sheets.count_rows(EP_LOG_RANGE, 3)

            # Set up a JSON style body object for the
            # batch update in the append values function
            range_body_values = {
                'value_input_option': 'USER_ENTERED',
                'data': [
                    {
                        'majorDimension': 'ROWS',
                        'range': f'EP Log!B{row_count}:J{row_count + len(guild_list) - 1}',
                        'values': guild_list
                    },
                    {
                        'majorDimension': 'ROWS',
                        'range': f'EP Log!R{row_count}:R{row_count + len(guild_list) - 1}',
                        'values': point_list
                    }
                ]
            }

            try:
                # send the HTTP write request to the Sheets API
                sheets.append_values(range_body_values)

            except HttpError as error:
                print(f"An error occurred: {error}")
                return error

            self.clear_ep()

    def import_ep_data(self):
        # This function retrieves EP attendance data
        # from the user's log file and writes it to
        # the main EP sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        # Check for successful validation
        if self.validate_ep_scan():
            if self.get_ep_time() == 'Guild Meeting':
                guild_list = self.read_discord()
            else:
                # The ep_time entry widget contains the selected
                # time stamp as a single string; it needs to be
                # split to parse out the month [0], the date[1],
                # and the time[2]; then the read_log method is
                # called to execute the sheets read function
                stripped_time = self.get_ep_time().split()
                guild_list = self.read_log(stripped_time[0], stripped_time[1], stripped_time[2])

            # Write the read results to the sheet
            self.set_ep_sheet(guild_list)
            self.set_ep_grid()

    def read_discord(self):
        # api_url = "http://63.42.237.42:9000/api/Discord/voice-channel-mains/1155966760919511176"    # general
        api_url = "http://63.42.237.42:9000/api/Discord/voice-channel-mains/1180415001731792966"    # guild-meeting
        # api_url = "http://63.42.237.42:9000/api/Discord/voice-channel-mains/1155966760919511177"  # group 1
        # api_url = "http://63.42.237.42:9000/api/Discord/voice-channel-mains/1158172083839311984"    # raid 1
        response = requests.get(api_url)
        names = response.json()
        guild_list = []

        if len(names) > 0:
            for key in names:
                time = datetime.now()

                curr_month = time.strftime('%b')
                curr_date = time.strftime('%d')
                curr_time = time.strftime('%H:%M:%S')
                curr_year = time.strftime('%Y')
                curr_level = sheets.get_player_level(key)
                curr_race = sheets.get_player_race(key)
                curr_name = key
                curr_class = sheets.get_player_class(key)

                curr_member = [curr_month, curr_date, curr_time, curr_year, curr_level, curr_class, curr_name]

                if (curr_race == 'High Elf'
                        or curr_race == 'Dark Elf'
                        or curr_race == 'Wood Elf'
                        or curr_race == 'Half Elf'):
                    curr_race = curr_race.split()
                    curr_member.append(curr_race[0])
                    curr_member.append(curr_race[1])
                else:
                    curr_member.append(curr_race)
                    curr_member.append('')

                guild_list.append(curr_member)

                self.set_ep_type('Meeting')
        else:
            display_error(NO_VOICE_ERROR)

        return guild_list

    @staticmethod
    def read_log(read_month, read_date, read_time):
        # This function runs an algorithm to locate
        # log file lines that are part of a /who guild
        # in game command, and formats them for importing
        # into the ep sheet
        # Parameters: read_month, read_data, read_time: all strings
        # Return: guild: list

        guild = []
        guild_tag = "<Seekers of Souls>"  # The tag to look for on each line of the log
        found_line = False  # A flag to indicate if an appropriate line has been found

        # Read the log file from the bottom up to save time
        with (FileReadBackwards(config_tab.get_log_file(), encoding="utf-8") as frb):
            for line in frb:
                # If the guild_tag is on the line, plus the month/date/time match what
                # the user specified, then a match is found """
                if guild_tag in line:
                    if read_month in line and read_date in line and read_time in line:
                        found_line = True
                        # Remove/replace unwanted formatting from EQ from the line
                        stripped_text = (str(line)
                                         .replace('[', '')
                                         .replace(']', '')
                                         .replace('(', '')
                                         .replace(')', '')
                                         .replace('Shadow Knight ', 'Shadowknight ')
                                         .replace(' AFK ', '')
                                         .replace('ANONYMOUS', 'ANONYMOUS ANON')
                                         .replace('Mon ', '')
                                         .replace('Tue ', '')
                                         .replace('Wed ', '')
                                         .replace('Thu ', '')
                                         .replace('Fri ', '')
                                         .replace('Sat ', '')
                                         .replace('Sun ', ''))

                        # Eliminate everything after the first < in the guild tab,
                        # then split the line into individual pieces to match the
                        # ep sheet and the EPGP log
                        sep = "<"
                        removed_end_text = stripped_text.split(sep, 1)[0]
                        guild.append(removed_end_text.split())

                    else:
                        # If found_line is already true but did not find guild tag and time
                        # stamp on the line, then we have reached the end of the /who guild
                        # output, so break out of the loop and return the guild list
                        if found_line:
                            break

        return guild

    def look_for_raids(self, initial):
        # This function sets up the algorithm in find_times
        # by creating the raid list and validating that a log file
        # exists. After running the algorithm it either notifies
        # user that no raid was found or populates the time stamps
        # combobox with the time stamp values
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        available_raids = []
        # grab current list of time stamps, to compare below
        curr_raid_count = self.get_ep_combo_values()
        self.set_ep_combo_values(available_raids)

        if config_tab.get_log_file():
            available_raids = self.find_times()

            if not available_raids:
                tkinter.messagebox.showinfo('No Raids Found',
                                            RAID_FOUND_ERROR)
                self.set_ep_combo_values(['Guild Meeting'])
            else:
                # If the user has clicked 'Refresh Raids' this
                # will trigger
                if not initial:
                    new_stamp = 0

                    # loop through find_times results
                    for new_item in available_raids:
                        found_item = False

                        # compare with original list
                        for curr_item in curr_raid_count:
                            # flag changes and break out of loop
                            if curr_item == new_item:
                                found_item = True
                                break

                        # increment counter if change found
                        if found_item is False:
                            new_stamp += 1

                    # finally, report differences to user
                    raid_diff = new_stamp

                    ttk.dialogs.Messagebox.show_info(
                        f"[ {raid_diff} ] additional time stamps found.",
                        "Raids Found", alert=True)

                self.set_ep_time('')
                # Add a combobox value for guild meetings
                available_raids.append('Guild Meeting')
                self.set_ep_combo_values(available_raids)

    @staticmethod
    def find_times():
        # This function runs a different algorithm to scan
        # user's log file and find instances of /who guild
        # that match raid parameters (i.e., not a /who all guild,
        # at least 12 players with a guild tag in the same zone)
        # and then save the associated time stamps. Algorithm runs
        # until 10 time stamps have been found or 200k log file lines
        # have been scanned, whichever occurs first
        # Parameters: none
        # Return: found_stamps: list

        player_log = config_tab.get_log_file()
        # Multiple strings for searching for the phrase
        # "There are <num> players in", but not including
        # the word "EverQuest"; that would indicate a
        # /who all, which we don't want
        there_tag = "There are"
        players_tag = "players in"
        everquest_tag = "EverQuest"
        # Guild tag, as in previous algorithm
        guild_tag = "<Seekers of Souls>"
        # Special string to account for roleplay players;
        # If a player is on /anon their guild tag will
        # not show up; to prevent the algorithm from thinking
        # the end of the /who guild command has been reached
        # we count the number of consecutive anon tags combined
        # with no guild tags; if more than 2 are found we assume
        # the end of the command has been found and exit
        # NOTE: it is possible that 3 guild members in a row are
        # /anon rather than /roleplay; if this is the case the
        # algorithm will fail to count this as a /who guild time
        # stamp; I think this possibility is remote, and I made
        # the decision to leave the cutoff point at 2
        anon_tag = "ANONYMOUS"
        # Initialize all these counting, boolean, and list variables
        count_started = False
        found_stamps = []
        num_players = 0
        num_stamps = 0
        num_lines = 0

        with (FileReadBackwards(player_log, encoding="utf-8") as frb):
            while True:
                line = frb.readline()

                # Use this to shut down the algorithm if 200k lines have been read
                num_lines += 1

                # If we have not started counting yet...
                if count_started is False:
                    # If "There are" and "players in" exist but not "EverQuest"
                    if there_tag in line and players_tag in line and everquest_tag not in line:
                        # Prepare a string to catch the zone the /who guild was executed in
                        zone = ''
                        # Split the line apart, delimiting by spaces
                        separated_text = line.split()
                        # Isolate the month, date, and time values
                        times = [separated_text[1], separated_text[2], separated_text[3]]
                        # Remove the unwanted log file formatting
                        times = (str(times).replace('[', '')
                                 .replace(']', '')
                                 .replace(',', '')
                                 .replace('\'', ''))
                        # Grab the player count from the "players in" line
                        player_count = separated_text[7]

                        # Make sure this isn't a bogus /who guild command, as
                        # indicated by "no players in..."; also make sure
                        # the player count is 12 or more
                        if player_count != 'no' and int(player_count) > 11:
                            # Zones can be as many as 3 or 4 separate words; to
                            # account for this we make a list of all the words
                            # starting after "players in"...
                            zone_parts = []

                            # Append them together into a single list...
                            for i in range(10, len(separated_text)):
                                zone_parts.append(separated_text[i])

                            # Turn it into a string...
                            zone = ' '.join(zone_parts)
                            # And get rid of an unwanted period at the end of the line
                            zone = zone.replace('.', '')
                            # All the setup pieces have been assembled, now start the
                            # count beginning with the next line above this one
                            count_started = True
                # If we have started counting, we take this branch
                else:
                    # If guild or anon (to account for players who went anon
                    # instead of roleplay) is found, increment players
                    if guild_tag in line or anon_tag in line:
                        num_players += 1
                    # If neither guild nor anon is found, we assume the
                    # end of the who request has been reached, so...
                    else:
                        # Check for enough players for a raid; if true
                        # append to time stamps, reset, and keep looking
                        if num_players > 11:
                            found_stamps.append(str(times) + '  |  ' + str(zone))
                            num_stamps += 1
                            times = ''
                            zone = ''
                            num_players = 0
                            # num_anon = 0
                            count_started = False
                        # It not enough players present, reset and keep looking
                        else:
                            times = ''
                            zone = ''
                            num_players = 0
                            # num_anon = 0
                            # num_garbage = 0
                            count_started = False

                # If we've already found 10 stamps or exceeded the line threshold
                # exit the loop and return the list of time stamps
                if num_stamps >= 10 or num_lines > 200000:
                    break

        return found_stamps

    def add_record(self):
        # This function opens the Add Record toplevel form
        # so the user can insert rows as needed
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        AddRecord(self, [])

    def validate_ep_entry(self):
        # This function provides the primary validation for the
        # ep tab
        # Parameters: self (inherit from TabEP parent)
        # Return: boolean

        sheet_values = self.get_ep_sheet()

        # If there is not at least one line of data return false
        if len(sheet_values) == 0:
            display_error(EMPTY_EP_ERROR)
            return False
        else:
            # If less than 8 items are in the first line return
            # false; this simulates the user clicking
            # 'manual entry' and then entering nothing
            if len(sheet_values[0][:15]) < 8:
                display_error(EMPTY_EP_ERROR)
                return False
            # If the user has not selected an EP type return false
            elif self._ep_type.get() == '':
                display_error(TYPE_EP_ERROR)
                return False
            # If all those boxes have been ticked, now we verify
            # that all player names entered match entries on the
            # Totals tab of the EPGP Log
            else:
                for row in sheet_values:
                    # Set the name found flag to false for each row
                    found_name = False

                    # Loop through the player list
                    for name in sheets.get_player_list():
                        # If a match is found, flag it True then exit
                        # the loop
                        if name == row[6]:  # noqa
                            found_name = True
                            break
                    # If a match was never found, display error and
                    # return false
                    if found_name is False:
                        display_error('[ ' + row[6] + SPELLING_EP_ERROR)  # noqa
                        return False

        # If no false returns have occurred, then exit with true
        return True

    def validate_ep_scan(self):
        # This function is a secondary validation, specifically
        # for the scan of the user's log file
        # Parameters: self (inherit from TabEP parent)
        # Return: boolean

        time_entered = self.get_ep_time()
        # Check to verify that user has entered something
        # in the time stamps combobox; if not return false
        if time_entered == '':
            display_error(SCAN_ERROR)
            return False
        # Check that user entry matches one of the available
        # time stamps; if not return false
        else:
            time_list = self.get_ep_combo_values()

            for stamp in time_list:
                # If a match is found function exits, thus
                # false return is never reached
                if stamp == time_entered:
                    return True

            # If code logic makes it this far then time stamp
            # match was never found; return false
            display_error(SCAN_ERROR)
            return False


class TabGP(ttk.Frame):
    # This class creates the GP Tab, responsible for:
    # - Mimicking the 'Get Priority' tab of the EPGP log,
    #   consisting of entering character names and loot bid
    #   levels, then writing the data to the Log to determine
    #   the winner
    # - Entering loot items won and writing the data to the GP Log tab

    def __init__(self, parent):
        super().__init__(parent)

        # ----- class members -----
        self._gp_date = tk.StringVar()
        self._gp_name = tk.StringVar()
        self._gp_loot = tk.StringVar()
        self._gp_loot_entry = AutocompleteCombobox(self, sheets.get_loot_names())
        self._gp_level = tk.StringVar()
        self._pr_winner = tk.StringVar()
        self._pr_ratio = tk.StringVar()
        self._pr_gear = tk.StringVar()
        # Custom class to create the scrolling canvas
        # frame that contains the get priority bidders
        self._list_frame = self.create_list_frame()

        self.configure(width=880, height=595)
        self.place(x=0, y=0)

        self.create_widgets()
        # convert date format to match EPGP Log
        self.set_gp_date(datetime.now().strftime("%m/%d/%y"))
        # configure notebook to disable mouse scrolling when not on
        # gp tab, this prevents list frame from being scrolled while
        # user is interacting with ep sheet
        parent.bind("<<NotebookTabChanged>>", self._list_frame.on_tab_selected)

    # ----- getters -----
    def get_gp_date(self):
        return self._gp_date.get()

    def get_gp_name(self):
        return self._gp_name.get()

    def get_gp_loot(self):
        return self._gp_loot.get()

    def get_gp_level(self):
        return self._gp_level.get()

    def get_pr_winner(self):
        return self._pr_winner.get()

    def get_pr_ratio(self):
        return self._pr_ratio.get()

    def get_pr_gear(self):
        return self._pr_gear.get()

    def get_list_frame(self):
        # This function returns the items view
        # object of the cells dictionary which
        # is created in the ListFrame class
        # Items contains key, value pairs
        return self._list_frame.get_cells().items()

    # ----- setters -----
    def set_gp_date(self, gp_value):
        self._gp_date.set(gp_value)

    def set_gp_name(self, gp_value):
        self._gp_name.set(gp_value)

    def set_gp_loot(self, gp_value):
        self._gp_loot.set(gp_value)

    def set_gp_loot_entry(self, gp_value):
        self._gp_loot_entry['values'] = gp_value

    def set_gp_level(self, gp_value):
        self._gp_level.set(gp_value)

    def set_pr_winner(self, gp_value):
        self._pr_winner.set(gp_value)

    def set_pr_ratio(self, gp_value):
        self._pr_ratio.set(gp_value)

    def set_pr_gear(self, gp_value):
        self._pr_gear.set(gp_value)

    # ----- widget setup -----
    def create_list_frame(self):
        # This function creates the get priority list frame
        # It reads the number of rows on the Get Priority tab
        # of the EPGP Log and creates the same number of
        # rows here
        # Parameters: self (inherit from TabGP parent)
        # Return: list_frame - instance of ListFrame class
        priority_frame = ttk.Frame(self)
        list_frame = ListFrame(priority_frame, sheets.get_priority_rows())
        priority_frame.place(x=20, y=80, width=540, height=250)

        return list_frame

    def create_widgets(self):
        # This class creates and places all the widgets
        # on the GP tab
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        priority_title = ttk.Label(self, text='Get Priority', font='Calibri 24 bold')
        priority_title.place(x=380, y=0)

        character_header = ttk.Label(self, text='Character')
        character_header.place(x=65, y=45)

        gear_level_header = ttk.Label(self, text='Gear Level')
        gear_level_header.place(x=230, y=45)

        priority_header = ttk.Label(self, text='Priority')
        priority_header.place(x=395, y=45)

        priority_winner_label = ttk.Label(self, text='Winner')
        priority_winner_label.place(x=620, y=60, width=150, height=40)

        priority_ratio_label = ttk.Label(self, text='Priority')
        priority_ratio_label.place(x=620, y=135, width=90, height=40)

        priority_level_label = ttk.Label(self, text='Gear Level')
        priority_level_label.place(x=620, y=210, width=150, height=40)

        priority_winner_text = ttk.Entry(self, font='Calibri 12', state=DISABLED, foreground="#A9BDBD",
                                         textvariable=self._pr_winner)
        priority_winner_text.place(x=620, y=95, width=150, height=40)

        priority_ratio_text = ttk.Entry(self, font='Calibri 12', state=DISABLED, foreground="#A9BDBD",
                                        textvariable=self._pr_ratio)
        priority_ratio_text.place(x=620, y=170, width=90, height=40)

        priority_level_text = ttk.Entry(self, font='Calibri 12', state=DISABLED, foreground="#A9BDBD",
                                        textvariable=self._pr_gear)
        priority_level_text.place(x=620, y=245, width=150, height=40)

        clear_bidders_button = ttk.Button(self, text='Clear', style='primary.Outline.TButton',
                                          command=self.clear_bidders)
        clear_bidders_button.place(x=420, y=350, width=140)

        find_winner_button = ttk.Button(self, text='Priority', style='primary.Outline.TButton',
                                        command=self.get_loot_winner)
        find_winner_button.place(x=570, y=350, width=140)

        copy_winner_button = ttk.Button(self, text='Copy', style='primary.Outline.TButton',
                                        command=self.copy_winner)
        copy_winner_button.place(x=720, y=350, width=140)

        # ------------------- GP / Priority separator -------------------
        gp_separator = ttk.Separator(self, orient='horizontal')
        gp_separator.place(x=10, y=400, width=855)

        # ------------------- GP Entry Section -------------------
        gp_title = ttk.Label(self, text='GP Entry', font='Calibri 24 bold')
        gp_title.place(x=327, y=410)

        gp_date_label = ttk.Label(self, text='Date')
        gp_date_label.place(x=20, y=450, width=100, height=30)

        gp_date_info_label = ttk.Label(self, text='click to change', font='Calibri 10')
        gp_date_info_label.place(x=25, y=515)

        gp_character_label = ttk.Label(self, text='Character')
        gp_character_label.place(x=155, y=450, width=160, height=30)

        gp_loot_label = ttk.Label(self, text='Loot')
        gp_loot_label.place(x=330, y=450, width=315, height=30)

        gp_level_label = ttk.Label(self, text='Gear Level')
        gp_level_label.place(x=660, y=450, width=200, height=30)

        gp_date_entry = ttk.Entry(self, font='Calibri 12', textvariable=self._gp_date)
        gp_date_entry.bind("<Button-1>", self.grab_date)
        gp_date_entry.place(x=20, y=480, width=120)

        gp_character_entry = AutocompleteCombobox(self, sheets.get_player_list())
        gp_character_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_name)
        gp_character_entry.place(x=155, y=480, width=160)

        self._gp_loot_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_loot)
        self._gp_loot_entry.place(x=330, y=480, width=315)

        gp_level_entry = AutocompleteCombobox(self, sheets.get_gear_points())
        gp_level_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_level)
        gp_level_entry.place(x=660, y=480, width=200)

        clear_gp_button = ttk.Button(self, text='Clear', style='primary.Outline.TButton',
                                     command=self.clear_gp)
        clear_gp_button.place(x=570, y=535, width=140)

        save_gp_button = ttk.Button(self, text='Save', style='primary.Outline.TButton',
                                    command=self.insert_gp)
        save_gp_button.place(x=720, y=535, width=140)

    # def print_frame(self):
    #     for key, value in self.get_list_frame():
    #         if str(key[1:]) == '0':
    #             print(key, value[0].get())

    # ----- class functions -----
    def grab_date(self, event):
        # This class displays the Date Picker widget
        # to the user in response to a left click on
        # the gp_date_entry widget
        # Parameters: self (inherit from TabGP parent)
        #             event - unused, but is auto-generated
        # Return: none

        cal = ttk.dialogs.DatePickerDialog()
        selected = cal.date_selected.strftime("%m/%d/%y")
        self.set_gp_date(selected)

    def clear_gp(self):
        # This class clears out the entry widgets
        # at the bottom of the tab, deliberately
        # leaving the date in place because typically
        # the user will be entering multiple items
        # per raid
        # Parameters:  self (inherit from TabGP parent)
        # Return: none

        self.set_gp_name('')
        self.set_gp_loot('')
        self.set_gp_level('')

    def clear_bidders(self):
        # This function clears out the get priority
        # scrolling canvas but resets the bid level
        # combobox to 'Major Upgrade', because that
        # is the bid level used most of the time
        # The function also clears the three winner
        # boxes
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        for key, value in self.get_list_frame():
            # print(str(key)[:1])
            match str(key)[:1]:
                case '1':
                    value[0].set('')
                case '2':
                    value[0].set('Major Upgrade')
                case '3':
                    value[0].set('')

        self.set_pr_winner('')
        self.set_pr_ratio('')
        self.set_pr_gear('')

    def insert_gp(self):
        # This function appends gp data (date, character,
        # loot, bid level) to the bottom of the GP Log
        # tab in the EPGP Log
        # Parameters: self (inherit from TabGP parent)
        # Return: possible Error

        # Check for successful GP validation
        if self.validate_gp():
            # Extra check to encourage matching item
            # names with what is already in the log
            if self.look_for_duplicates():

                # Create the list that will be inserted
                values = [
                    [
                        self.get_gp_date(),
                        self.get_gp_name(),
                        self.get_gp_loot(),
                        self.get_gp_level()
                    ]
                ]

                # Obtain the insertion point for the GP LOg;
                # i.e., the first empty cell in column C
                row_count = sheets.count_rows(GP_LOG_RANGE, 2)

                # Set up a JSON style body object for the
                # batch update in the append values function
                value_range = {
                    'value_input_option': 'USER_ENTERED',
                    'data': {
                        'majorDimension': 'ROWS',
                        'range': f'GP Log!C{row_count}:F{row_count}',
                        'values': values
                    }
                }

                try:
                    # Send the HTTP update request
                    sheets.append_values(value_range)

                except HttpError as error:
                    print(f"An error occurred: {error}")
                    return error

                # Clear out the GP and Get Priority data
                self.clear_gp()
                self.clear_bidders()
        # The only time an error pop-up is displayed
        # outside a validation function
        else:
            display_error(ENTER_GP_ERROR)

    def look_for_duplicates(self):
        # This function checks for a matching item name
        # entry in the GP Log tab; this is to try to
        # enforce correct spelling and consistent item
        # naming conventions in the log
        # Parameters: self (inherit from TabGP parent)
        # Return: boolean

        item_entry = self.get_gp_loot()
        # Set the found flag to false by default,
        # so that only a match gives a true result
        item_found = False

        # Loop through the item list, which was
        # already read from the EPGP Log at startup
        for item in sheets.get_loot_names():
            if item == item_entry:
                item_found = True
                break

        # If the item was never spelled, prompt user
        # to confirm they are entering the correct
        # item
        if item_found is False:
            confirm = ttk.dialogs.Messagebox.yesno(
                f"[ {item_entry} ] does not appear in the EPGP Log.\n"
                "Are you sure you want to enter it?",
                "Confirm Item Insertion", alert=True)

            # If the user clicks yes, append the new item
            # to the loot names list and return true
            if confirm == 'Yes':
                sheets.update_loot_names(item_entry)
                self.set_gp_loot_entry(sheets.get_loot_names())
                return True
            # Otherwise, return false
            else:
                return False
        else:
            return True

    def write_loot_contestants(self):
        # This function clears out the cells on the
        # Get Priority tab, then writes in the names
        # and bid levels entered in the list frame
        # Parameters: self (inherit from TabGP parent)
        # Return: possible Error

        sheets.clear_values()

        items = []
        values = []

        # As described earlier, get_list_frame returns
        # a key, value paired items view object; we
        # are using that object to create a 2D list
        for key, value in self.get_list_frame():
            # Each 'row' in items contains a char name
            # and a bid level; the case here is the
            # column, so for the first two columns
            # we append the data to what will become
            # the inner list of the current outer list
            # index. Then when the third column is reached
            # (which is the priority number box (not used
            # yet) we append the inner list to the outer
            # and reset the inner list.
            match str(key)[:1]:
                case '1':
                    items.append(value[0].get())
                case '2':
                    items.append(value[0].get())
                case '3':
                    values.append(items)
                    items = []
        # Grab the length of the 2d list...
        upper_bound = len(values)

        # And use it to 'None' out the bid level
        # from the middle column if the user has
        # for some reason skipped rows in the list
        # frame; this is a fail-safe to ensure that
        # only rows with good data are written to the
        # Get Priority tab
        for i in range(upper_bound):
            if values[i][0] == '':
                values[i][0] = None
                values[i][1] = None

        # Create the range string for the JSON body
        body_range = f'Get Priority!B3:C{len(values) + 2}'

        # Set up a JSON style body object for the
        # batch update in the append values function
        range_body_values = {
            'value_input_option': 'USER_ENTERED',
            'data': {
                'majorDimension': 'ROWS',
                'range': body_range,
                'values': values
            }
        }

        try:
            # Send the HTTP request to the EPGP Log
            sheets.append_values(range_body_values)

        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def read_loot_results(self):
        # This function reads in two lists from the Get
        # Priority tab. The first is the winning bidder,
        # their bid level, and their priority ratio. The
        # second list is the ratios of all bidders. This
        # data is then plugged into the app window
        # Parameters: self (inherit from TabGP parent)
        # Return: possible Error

        try:
            winner = sheets.read_values(LOOT_WINNER_RANGE, 'FORMATTED_VALUE')
            gl_pr = sheets.read_values(GET_PRIORITY_RANGE, 'FORMATTED_VALUE')
            num_bidders = len(gl_pr)
            index = 0

            for key, value in self.get_list_frame():
                # If the column is 3...
                if str(key[:1]) == '3':
                    # And if the current index is less
                    # than the len of the ratios pulled
                    # from the log (so we know when to
                    # stop), then set the current items
                    # index to the ratio value that was
                    # read in; this updates the live widget
                    # We can do this because
                    # the get_list_frame function is
                    # returning an updatable object
                    # of the list frame contents
                    if index < num_bidders:
                        value[0].set(gl_pr[index])
                        index += 1
                    else:
                        break

            # If #N/A is read in from the Log, this indicates
            # an error, so notify the user; otherwise, write
            # the winning data to the window
            if winner[0][0] != '#N/A':
                self.set_pr_winner(winner[0][1])
                self.set_pr_ratio(winner[0][0])
                self.set_pr_gear(winner[0][2])
            else:
                display_error(READ_WINNER_ERROR)

        except HttpError as err:
            print(err)

    def get_loot_winner(self):
        # This function wraps the get priority
        # writing/reading functions and provides
        # validation
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        if self.validate_priority():
            self.write_loot_contestants()
            self.read_loot_results()
        else:
            display_error(FIND_WINNER_ERROR)

    def copy_winner(self):
        # This function wraps the get priority
        # writing/reading functions and provides
        # validation
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        # Basically, if there's no winner validation fails
        if self.validate_copy_winner():
            self.set_gp_name(self.get_pr_winner())
            self.set_gp_level(self.get_pr_gear())
        else:
            display_error(COPY_WINNER_ERROR)

    def validate_gp(self):
        # This function provides validation for the
        # insert_gp function
        # Parameters: self (inherit from TabGP parent)
        # Return: boolean

        # Simple validation; if any field is empty
        # return false, otherwise true
        if (self.get_gp_date() == ''
                or self.get_gp_name() == ''
                or self.get_gp_loot() == ''
                or self.get_gp_level() == ''):
            return False
        else:
            return True

    def validate_priority(self):
        # This function provides validation for the
        # get priority function
        # Parameters: self (inherit from TabGP parent)
        # Return: boolean

        num_matches = 0
        char_found = False
        level_found = False

        # Grab the list frame contents and loop through it
        for key, value in self.get_list_frame():
            # If we are on column 1, grab the name
            if str(key[:1]) == '1':
                character = value[0].get()

                # Then compare it with the player list
                for item in sheets.get_player_list():
                    # If a match is found set the flag to true
                    if item == character:
                        char_found = True

            # Run the same logic for the bid levels
            if str(key[:1]) == '2':
                gear_level = value[0].get()

                for item in sheets.get_bid_levels():
                    if item == gear_level:
                        level_found = True

            # If we reached the third column and found a match
            # for both column 1 and 2 then increment num_matches
            if str(key[:1]) == '3' and char_found is True and level_found is True:
                num_matches += 1

            # Every time we hit the third column, reset the match
            # flags to 0; this ensures that matches for each row
            # are kept separate from each other
            if str(key[:1]) == '3' and str(key[1:] != '0'):
                char_found = False
                level_found = False

        # If at least two char/level matches were found
        # then we can run Get Priority, so return true
        if num_matches >= 2:
            return True
        else:
            return False

    def validate_copy_winner(self):
        # This function provides validation for the
        # get priority function
        # Parameters: self (inherit from TabGP parent)
        # Return: boolean

        # Another simple one; is the winner
        # field empty or not
        if self.get_pr_winner() == '':
            return False
        else:
            return True


class TabConfig(ttk.Frame):
    # This class creates the Config Tab, responsible for:
    # - Retrieving/Changing the path to the user's EQ log file;
    #   this path is stored in 'config' in the same directory
    #   as the application
    # - Changing and saving the current window theme

    def __init__(self, parent):
        super().__init__(parent)

        # ----- class members -----
        self._log_file = ttk.StringVar()
        self._app_theme = ttk.StringVar()

        self.configure(width=880, height=595)
        self.place(x=0, y=0)

        self.create_widgets()
        # Read in the log file path at startup
        self.read_log_file()

    # ----- getters -----
    def get_log_file(self):
        return self._log_file.get()

    def get_app_theme(self):
        return self._app_theme.get()

    # ----- setters -----
    def set_log_file(self, config_value):
        self._log_file.set(config_value)

    def set_app_theme(self, config_value):
        self._app_theme.set(config_value)

    # ----- widget setup -----
    def create_widgets(self):
        log_label = ttk.Label(self, text='Log File')
        log_label.place(x=20, y=20, width=100, height=30)

        log_entry = ttk.Label(self, font='Calibri 12', textvariable=self._log_file, borderwidth=10, relief=tk.SUNKEN)
        log_entry.place(x=20, y=50, width=650, height=40)

        change_log_button = ttk.Button(self, text='Change Log', style='primary.Outline.TButton',
                                       command=lambda: self.change_log_file(False))
        change_log_button.place(x=720, y=122, width=140)

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

    def change_log_file(self, initial):
        # This function allows users to set their
        # log file path. Typically, this should
        # only have to be done once
        # Parameters:  self (inherit from TabConfig parent)
        #              initial - boolean
        # Return: boolean

        # Pop an open file dialog
        file_path = filedialog.askopenfilename(title="Select Log File", filetypes=[("Text files", "*.txt")])

        # Set the log file, save it, then run
        # start-up functions
        if file_path:
            self._log_file.set(file_path)
            self.save_log_file()

            # Have to check whether this is being run
            # for the first time ever; if it is then the TabConfig
            # class is not finished instantiating, and will
            # throw an error when the look_for_raids method
            # is called, because that method depends on the
            # get_log_file getter method of this class
            if initial is False:
                ep_tab.look_for_raids(False)
                ep_tab.clear_ep()

            return True
        else:
            return False

    def save_log_file(self):
        # This function allows users to set their
        # log file path. Typically, this should
        # only have to be done once
        # Parameters:  self (inherit from TabConfig parent)
        # Return: none

        path_to_file = "config"

        with open(path_to_file, 'w') as file:
            content = self.get_log_file()
            file.write(content)

    def read_log_file(self):
        # This function is run at start-up to read in
        # the path to the player's EQ log file
        # Parameters: self (inherit from TabConfig parent)
        # Return: player_log (string)

        path_to_file = "config"

        # If the 'config' file path does not exist display
        # error and run the change_log function
        if not os.path.isfile(path_to_file):
            app.overrideredirect(True)
            app.attributes('-alpha', 0)
            app.deiconify()

            display_error(NO_LOG_ERROR)

            self.change_log_file(True)

        # Now check for the 'config' file again; if it still
        # doesn't exist (because user didn't choose a file
        # or something else weird is going on, display
        # an error and shut down the app
        if not os.path.isfile(path_to_file):
            display_error(STUBBORN_LOG_ERROR)
            sys.exit()
        else:
            # If the config file is found, read in the path
            # to the player's log file, set the label widget
            # to it, then return the path as a string
            with open(path_to_file, "r") as file:
                line = file.readline()
                self.set_log_file(line)
                player_log = line

        return player_log


class ListFrame(ttk.Frame):
    # A Tkinter scrollable canvas with internal widgets.
    # Created by Christian Koch; original source code at
    # https://github.com/clear-code-projects/tkinter-complete
    # Modified by Timothy Wise for use with Seekers of Souls' EPGP loot
    # system on the Project Quarm emulated Everquest server

    def __init__(self, parent, num_rows):
        super().__init__(master=parent)
        self.pack(expand=True, fill='both')

        # ----- widget data -----
        self._num_rows = num_rows

        # Each row is 35 pixels high on windows 10, 40 on 11
        if platform.platform()[:10] == 'Windows-10':
            self._list_height = self._num_rows * 35
        else:
            self._list_height = self._num_rows * 40

        self._cells = {}

        # ----- canvas -----
        self.canvas = tk.Canvas(self, background='red', scrollregion=(0, 0, self.winfo_width(), self._list_height))
        self.canvas.pack(expand=True, fill='both')

        # ----- display frame -----
        self.frame = ttk.Frame(self)

        # Create a number of rows equal to the number
        # on the Get Priority tab of the EPGP Log
        for index in range(self._num_rows):
            self.create_item(index).pack(expand=True, fill='both', pady=1, padx=1)

        # ----- scrollbar -----
        self.scrollbar = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind_all('<MouseWheel>',
                             lambda event: self.canvas.yview_scroll(-int(event.delta / 60), "units"))
        self.scrollbar.place(relx=1, rely=0, relheight=1, anchor='ne')

        # ----- events -----
        # The enter and leave events ensure that the mouse wheel
        # only scrolls when the cursor is on the list frame; This
        # ensures that the list frame is not scrolled when the user
        # is scrolling a combobox elsewhere in the window
        self.frame.bind('<Enter>', self.bind_to_mousewheel)
        self.frame.bind('<Leave>', self.unbind_from_mousewheel)
        self.bind('<Configure>', self.update_size)

    def get_cells(self):
        # This function returns the cells dictionary
        # to the caller
        # Parameters: self (inherit from ListFrame parent)
        # Return: _cells (dictionary)

        return self._cells

    def update_size(self, event):
        # This function adjusts the size of the canvas to
        # the number of rows being rendered
        # Parameters: self (inherit from ListFrame parent)
        #             event - unused, but is auto-generated
        # Return: none

        height = self._list_height

        self.canvas.create_window(
            (0, 0),
            window=self.frame,
            anchor='nw',
            width=self.winfo_width(),
            height=height)

    def on_tab_selected(self, event):
        # This function turns off the mousewheel for the canvas
        # if the GP Tab is not selected; this prevents the frame
        # from being scrolled when the user is on another tab
        # Parameters: self (inherit from ListFrame parent)
        #             event - the widget generating the event
        # Return: none

        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")

        if tab_text == "Gear Points":
            self.bind_to_mousewheel(event)
        else:
            self.unbind_from_mousewheel(event)

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 60)), "units")

    def bind_to_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

    def unbind_from_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def create_item(self, index):
        # This function creates a copy of each list frame
        # row up to the number of rows read from the GP Log
        # Parameters: self (inherit from ListFrame parent)
        #             index - int
        # Return: frame

        frame = ttk.Frame(self.frame)

        # ----- grid layout -----
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1, uniform='a')
        frame.columnconfigure(1, weight=4, uniform='a')
        frame.columnconfigure(2, weight=4, uniform='a')
        frame.columnconfigure(3, weight=3, uniform='a')
        frame.columnconfigure(4, weight=1, uniform='a')

        # Generate a unique id for every cell in the grid,
        # consisting of the column number followed by the
        # row number
        player_id = f'{1}{index}'
        gear_id = f'{2}{index}'
        ratio_id = f'{3}{index}'

        # Set up the tkinter variables to hold the values
        # for later insertion into the frame list
        player_var = StringVar(frame, '', player_id)
        gear_var = StringVar(frame, '', gear_id)
        ratio_var = StringVar(frame, '', ratio_id)

        # widgets
        count_label = ttk.Entry(frame, font="Calibri 12", foreground="#ffffff")
        count_label.insert(0, f' {index + 1}')
        count_label.configure(state=DISABLED)
        count_label.grid(row=0, column=0, sticky='nsew')

        player_combo = AutocompleteCombobox(frame, sheets.get_player_list())
        player_combo.configure(font=('Calibri', 12), height=40, textvariable=player_var)
        player_combo.grid(row=0, column=1, sticky='nsew')
        player_combo.unbind_class('TCombobox', '<MouseWheel>')

        gear_combo = AutocompleteCombobox(frame, sheets.get_bid_levels())
        gear_combo.configure(font=('Calibri', 12), height=40, textvariable=gear_var)
        gear_combo.grid(row=0, column=2, sticky='nsew')
        gear_combo.current(1)

        ratio_label = ttk.Entry(frame, font="Calibri 12", state=DISABLED, foreground="#A9BDBD",
                                textvariable=ratio_var)
        ratio_label.grid(row=0, column=3, sticky='nsew')

        # Insert each value into the 2d list at the
        # unique cell position set up earlier
        self._cells[player_id] = [player_var]
        self._cells[gear_id] = [gear_var]
        self._cells[ratio_id] = [ratio_var]

        return frame


# ---------- general classes ----------
class AutocompleteCombobox(ttk.Combobox):
    # A Tkinter widget that features autocompletion.
    # Created by Mitja Martini on 2008-11-29.
    # Updated by Russell Adams, 2011/01/24 to support Python 3 and Combobox.
    # Updated by Dominic Kexel to use Tkinter and ttk instead of tkinter and tkinter.ttk
    # Updated by Timothy Wise to follow Python class conventions for use with
    #     Seekers of Souls' EPGP loot system on the Project Quarm emulated Everquest server
    #    Licensed same as original (not specified?), or public domain, whichever is less restrictive.

    def __init__(self, parent, completion_list):
        super().__init__(master=parent)

        self._completion_list = completion_list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.configure(font=('Calibri', 12), height=40)

        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Setup our popup menu

    def autocomplete(self, delta=0):
        # autocomplete the Combobox, delta may be 0/1/-1 to cycle through possible hits
        if delta:  # need to delete selection otherwise we would fix the current position
            self.delete(self.position, tk.END)
        else:  # set position to end so selection starts where textentry ended
            self.position = len(self.get())
        # collect hits
        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):  # Match case insensitively
                _hits.append(element)
        # if we have a new hit list, keep this in mind
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        # only allow cycling if we are in a known hit list
        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)
        # now finally perform the auto-completion
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        global key_strokes

        # key_strokes += 1
        time.sleep(0.2)
        # if key_strokes >= 2:
        # event handler for the keyrelease event on this widget
        if event.keysym == "BackSpace":
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keysym == "Left":
            if self.position < self.index(tk.END):  # delete the selection
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1  # delete one character
                self.delete(self.position, tk.END)
        if event.keysym == "Right":
            self.position = self.index(tk.END)  # go to end (no selection)
        if len(event.keysym) == 1:
            self.autocomplete()
        # No need for up/down, we'll jump to the popup
        # list at the position of the autocompletion

            # key_strokes = 0


class GoogleSheets:
    def __init__(self):
        try:
            # class members
            self._service = self.get_service()

            self._priority_rows = 0
            self._player_list = []
            self._player_dict = {}
            self._effort_points = []
            self._gear_points = []
            self._bid_levels = []
            self._max_level = ''
            self._loot_names = []

        except HttpError as error:
            print(f"An error occurred: {error}")
            display_error("An HTTP error occurred. You may need to request access to EPGP-Interface."
                          "\nProgram terminating.")
            quit()

    # class getters
    def get_priority_rows(self):
        return self._priority_rows

    def get_player_list(self):
        return self._player_list

    def get_player_level(self, key):
        return self._player_dict[key]['level']

    def get_player_class(self, key):
        return self._player_dict[key]['class']

    def get_player_race(self, key):
        return self._player_dict[key]['race']

    def get_effort_points(self):
        return self._effort_points

    def get_gear_points(self):
        return self._gear_points

    def get_loot_names(self):
        return self._loot_names

    def get_max_level(self):
        return self._max_level

    def get_bid_levels(self):
        return self._bid_levels

    # class setters
    # Rather complicated looking member setters; but we're
    # just reading in value lists from the EPGP Log
    def set_priority_rows(self):
        # Get the len of the rows on the Get Priority tab;
        # We have to use FORMULA for the valuerenderoption here
        # so we can tell how many rows have been inserted

        self._priority_rows = (
            len(self.read_values(COUNT_PRIORITY_RANGE, 'FORMULA')))

    # All the rest of the setters just use FORMATTED_VALUE,
    # which just means read the contents of the cell rather
    # than the formula; we are also using the sum function
    # to flatten the 2D lists retrieved into 1D lists
    def set_player_list(self):
        self._player_list = (
            sum(self.read_values(RAIDER_LIST_RANGE, 'FORMATTED_VALUE'), []))

        self._player_dict = self.build_player_dict()

    def build_player_dict(self):
        # This function constructs a nested dictionary
        # with a top level key of every player name on
        # the EPGP Totals tab, and a nested dictionary
        # with keys race and class, for the purpose of
        # auto-populating the manual entry form
        # Parameters: self (inherit from GoogleSheets parent)
        # Return: master_dict (dictionary)

        master_dict = {}
        # read in all the values in the EPGP Log
        # from columns F, G, H, I, and J
        raw_levels = self.read_values("EP Log!F3:F", "FORMATTED_VALUE")
        raw_classes = self.read_values("EP Log!G3:G", "FORMATTED_VALUE")
        raw_names = self.read_values("EP Log!H3:H", 'FORMATTED_VALUE')
        raw_races = self.read_values("EP Log!I3:I", 'FORMATTED_VALUE')

        # Outer loop through Totals tab player list
        for player in self.get_player_list():
            # Start counter with max of list from column I, need to
            # use this column because it could be shortest
            counter = len(raw_races) - 1

            # Inner loop, decrementing through player names to find a match,
            # but using race list to avoid potential out of range errors
            for raw_race in reversed(raw_races):
                formatted_player = strip_garbage(raw_names[counter])

                # If player name match found
                if player == formatted_player:
                    # Only process further if this is a "valid"
                    # match with no ANON or Unknown values
                    if (raw_classes[counter] != "ANON"
                            and len(raw_races[counter]) > 0
                            and str(raw_races[counter]).find('Unknown') == -1):
                        # Combine race values in columns I and J, but
                        # remove any trailing white space if this is not
                        # and "Elf" race
                        player_level = (strip_garbage(raw_levels[counter]))

                        if player_level == 'ANONYMOUS':
                            player_level = sheets.get_max_level()

                        # print(raw_races_extra[counter])
                        player_race = strip_garbage(raw_race)

                        if (player_race == 'High'
                                or player_race == 'Dark'
                                or player_race == 'Wood'
                                or player_race == 'Half'):
                            player_race = player_race + ' Elf'

                        player_class = strip_garbage(raw_classes[counter])
                        # add the inner sub keys, values to the outer key
                        master_dict[player] = {'level': player_level, 'class': player_class, 'race': player_race}
                        break

                # Decrement the counter
                counter -= 1

        return master_dict

    def set_effort_points(self):
        self._effort_points = (
            sum(self.read_values(EFFORT_POINTS_RANGE, 'FORMATTED_VALUE'), []))

    def set_gear_points(self):
        self._gear_points = (
            sum(self.read_values(GEAR_POINTS_RANGE, 'FORMATTED_VALUE'), []))

    def set_bid_levels(self):
        self._bid_levels = (
            sum(self.read_values(BID_LEVELS_RANGE, 'FORMATTED_VALUE'), []))

    def set_max_level(self):
        # This function finds the maximum level listed on
        # the Totals tab of the EPGP Log, for autopopulating
        # the AddForm level field
        # Parameters: self (inherit from Sheets parent)
        # Return: curr_max (str)

        curr_max = 0
        # Use the sum function to flatten 2D list into 1D list
        level_list = sum(self.read_values(LEVEL_RANGE, 'FORMATTED_VALUE'), [])

        # Loop through results, ignoring anons
        # and recording max number found
        for level in level_list:
            if level != 'ANONYMOUS':
                if int(level) > curr_max:
                    curr_max = int(level)

        self._max_level = str(curr_max)

    def set_loot_names(self):
        # This one requires some special treatment; the GP Log
        # will contain multiple copies of many items, so we
        # want to remove the duplicates for our loot list,
        # and then sort it (the player list is already sorted
        # in the log and we want the other lists in the order
        # entered in the log
        self._loot_names = (
            list(set(sum(self.read_values(LOOT_NAMES_RANGE, 'FORMATTED_VALUE'), []))))
        self._loot_names.sort()

    def update_loot_names(self, sheets_value):
        # This setter function appends a new
        # loot item to the master item list,
        # then sorts the list
        self._loot_names.append(sheets_value)
        self._loot_names.sort()

    def get_service(self):
        # This function provides the authentication logic
        # for the Google Sheets API. The code here is
        # mostly boilerplate from Google, with a small
        # threading loop to account for the user not
        # confirming their Google account
        creds = None
        re_auth = False

        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except RefreshError:
                    tkinter.messagebox.showinfo('Authentication Expired',
                                                'Authentication token has expired.\n'
                                                'Please authenticate again.')
                    os.remove('token.json')

                    re_auth = True
            else:
                re_auth = True

            if re_auth is True:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build('sheets', 'v4', credentials=creds)

    def count_rows(self, row_range, offset):
        result = (
            self._service.spreadsheets().values()
            .get(spreadsheetId=SPREADSHEET_ID, range=row_range, valueRenderOption='FORMATTED_VALUE')
            .execute()
        )

        row_counter = 0

        for item in result['values']:
            row_counter += 1

        row_counter += offset

        return row_counter

    def append_values(self, range_body_values):
        # This function writes a 2d list of values
        # to the bottom of the range specified
        # Parameters: self (inherits from GoogleSheets parent)
        #             values - list
        #             append_range - constant in A1 notation
        # Return: result object indicating success or failure

        result = (
            self._service.spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=range_body_values)
            .execute()
        )

        # print(f"{(result.get('updates').get('updatedCells'))} cells appended.")

        return result

    # -------------------------------------------------------------
    # ----- Old append_values function that used Sheets ___________
    # ----- API append method; Caused problems with EP  -----------
    # ----- Log writes; leaving in the code for now     -----------
    # -------------------------------------------------------------

    # def append_values(self, values, append_range):
    #     # This function writes a 2d list of values
    #     # to the bottom of the range specified
    #     # Parameters: self (inherits from GoogleSheets parent)
    #     #             values - list
    #     #             append_range - constant in A1 notation
    #     # Return: result object indicating success or failure
    #
    #     body = {"values": values}
    #     result = (
    #         self._service.spreadsheets()
    #         .values()
    #         .append(
    #             spreadsheetId=SPREADSHEET_ID,
    #             range=append_range,
    #             # valueInputOption='RAW',
    #             valueInputOption='USER_ENTERED',
    #             body=body,
    #         )
    #         .execute()
    #     )
    #     print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
    #
    #     return result

    def clear_values(self):
        # This function clears the cell contents
        # of the Get Priority tab just prior to
        # writing in new characters and loot levels
        # Parameters: self (inherits from GoogleSheets parent)
        # Return: result object indicating success or failure

        values = []
        upper_bound = self._priority_rows

        for i in range(upper_bound):
            values.append(['', ''])

        body = {"values": values}

        result = (
            self._service.spreadsheets()
            .values()
            .update(
                spreadsheetId=SPREADSHEET_ID,
                range=CHARACTER_GEAR_RANGE,
                valueInputOption='USER_ENTERED',
                body=body
            )
            .execute()
        )

        # print(f"{(result.get('updates').get('updatedCells'))} cells cleared.")

        return result

    def read_values(self, read_range, value_render):
        # This function reads in a list of values
        # from the range specified
        # Parameters: self (inherits from GoogleSheets parent)
        #             read_range - constant in A1 notation
        #             value_render - constant, determines
        #               whether cell contents or formula is read
        # Return: values - 2d list

        result = (
            self._service.spreadsheets().values()
            .get(spreadsheetId=SPREADSHEET_ID, range=read_range, valueRenderOption=value_render)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return

        return values


# ---------- stand-alone functions ----------
def display_error(message):
    # This function serves as a template for all error
    # messages in the application
    ttk.dialogs.Messagebox.show_error(message, 'Missing Criteria')


def start_up():
    # This function runs several time-consuming
    # start-up tasks, mostly utilizing the sheets
    # object to communicate with the EPGP Log.
    # In addition, it handles the functionality
    # of the progress bar on the load screen
    # Parameters: none
    # Return: none

    # Ensure that the tabs are operating in the
    # global namespace so they can be referenced
    # in other parts of the code
    global ep_tab, gp_tab, config_tab

    # Set the sheets class members, updating
    # the load screen after each one
    load_screen.update_label('    Counting priority rows    ')
    sheets.set_priority_rows()
    load_screen.update_progress(12.5)

    load_screen.update_label('    Assembling player list    ')
    sheets.set_player_list()
    load_screen.update_progress(12.5)

    load_screen.update_label('   Obtaining EP categories    ')
    sheets.set_effort_points()
    load_screen.update_progress(12.5)

    load_screen.update_label('       Reading GP types       ')
    sheets.set_gear_points()
    load_screen.update_progress(12.5)

    load_screen.update_label('      Finding bid levels      ')
    sheets.set_bid_levels()
    load_screen.update_progress(12.5)

    load_screen.update_label('    Determining max level     ')
    sheets.set_max_level()
    load_screen.update_progress(12.5)

    load_screen.update_label('      Building loot list      ')
    sheets.set_loot_names()
    load_screen.update_progress(12.5)

    # Create the main tab classes
    ep_tab = TabEP(app)
    gp_tab = TabGP(app)
    config_tab = TabConfig(app)

    # Run an initial scan of the log file to
    # set the available raid time stamps
    load_screen.update_label(' Scanning for raid time stamps')
    ep_tab.look_for_raids(True)
    ep_tab.clear_ep()
    load_screen.update_progress(12.5)

    # Add the tab classes to the main window
    app.add_tab(ep_tab, 'Effort Points')
    app.add_tab(gp_tab, 'Gear Points')
    app.add_tab(config_tab, 'Configure')

    # Restore the main window (it was hidden immediately
    # on creation so the load window could be
    # displayed alone
    app.overrideredirect(False)
    app.attributes('-alpha', 100)
    app.deiconify()

    # Kill the load window, as its no longer needed
    load_screen.loading_window.destroy()


def strip_garbage(value):
    value = (str(value)
             .replace('[', '')
             .replace('\'', '')
             .replace(']', ''))

    return value


# ---------- main execution code ----------
if __name__ == "__main__":
    # global namespace for main window tabs
    global ep_tab, gp_tab, config_tab

    # global variable to slow down autocomplete
    key_strokes = 0

    # Create the root window
    app = Notebook()

    # Set up the Google Sheets API communication framework
    sheets = GoogleSheets()

    # Create loading screen and call it after a brief
    # pause to allow it to display, then run start-up
    load_screen = Load(app)
    load_screen.loading_window.after(100, start_up)

    app.run()
