import tkinter as tk
import ttkbootstrap as ttk
from datetime import datetime
import re

from classes.AutocompleteCombobox import AutocompleteCombobox
# from classes.GoogleSheets import GoogleSheets
from classes.Helper import Helper

class AddEP(tk.Toplevel):
    # This class creates the toplevel entry form
    # to add records manually to the ep sheet

    def __init__(self, parent, sheets, data_list):
        super().__init__(parent)

        # Obtain parent window position and place
        # this form relative to that
        parent_x_pos = parent.winfo_rootx()
        parent_y_pos = parent.winfo_rooty()
        geometry = '480x240+%d+%d' % (parent_x_pos + 150, parent_y_pos + 180)

        #
        self.transient(parent)
        self.grab_set()

        if len(data_list) < 1:
            self.title('Add Record')
        else:
            self.title('Edit Record')

        self.geometry(geometry)
        self.resizable(False, False)

        # Class data members
        self._ep_tab = parent
        self._sheets = sheets
        self._helper = Helper()

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

        # hard-coded lists
        self.MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        self.CLASSES = ['Bard', 'Beastlord', 'Cleric', 'Druid', 'Enchanter', 'Magician', 'Monk', 'Necromancer',
                   'Paladin', 'Ranger', 'Rogue', 'Shadow Knight', 'Shaman', 'Warrior', 'Wizard', 'Unknown']
        self.RACES = ['Barbarian', 'Dark Elf', 'Dwarf', 'Erudite', 'Froglok', 'Gnome', 'Half Elf', 'Halfling', 'High Elf',
                 'Human', 'Iksar', 'Ogre', 'Troll', 'Vah Shir', 'Wood Elf', 'Unknown']

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
            if len(self._ep_tab.get_ep_sheet()) > 0:
                self.set_add_month(self._ep_tab.get_ep_sheet_cell("A1"))
                self.set_add_date(self._ep_tab.get_ep_sheet_cell("B1"))
                self.set_add_time(self._ep_tab.get_ep_sheet_cell("C1"))
                self.set_add_year(self._ep_tab.get_ep_sheet_cell("D1"))
            else:
                # Use date/time formatting to make sure data
                # matches EPGP Log
                self.set_add_month(time.strftime('%b'))
                self.set_add_date(time.strftime('%d'))
                self.set_add_time(time.strftime('%H:%M:%S'))
                self.set_add_year(time.strftime('%Y'))

            # Just set a default level of 50; user will have
            # to modify if needed
            self.set_add_level(self._sheets.get_max_level())

    def create_widgets(self):
        # This function lays out the widgets on the form
        # Parameters: self (inherit from AddRecord parent)
        # Return: none

        self._entry_frame.pack(expand=True, fill='both')

        add_month_label = ttk.Label(self._entry_frame, text='Month', font='Calibri 14')
        add_month_label.place(x=10, y=10, width=70)

        add_month_entry = AutocompleteCombobox(self._entry_frame, self.MONTHS)
        add_month_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_month)
        add_month_entry.place(x=10, y=45, width=70)

        add_date_label = ttk.Label(self._entry_frame, text='Date', font='Calibri 14')
        add_date_label.place(x=90, y=10, width=50)

        dates = [str(x).zfill(2) for x in range(1, 32)]
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

        levels = [str(x).zfill(1) for x in range(1, 66)]
        add_level_entry = ttk.Combobox(self._entry_frame, values=levels)
        add_level_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_level)
        add_level_entry.place(x=330, y=45, width=50)

        add_class_label = ttk.Label(self._entry_frame, text='Class', font='Calibri 14')
        add_class_label.place(x=10, y=90, width=150)

        add_class_entry = AutocompleteCombobox(self._entry_frame, self.CLASSES)
        add_class_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_class)
        add_class_entry.place(x=10, y=125, width=150)

        add_name_label = ttk.Label(self._entry_frame, text='Name', font='Calibri 14')
        add_name_label.place(x=170, y=90, width=160)

        add_name_entry = AutocompleteCombobox(self._entry_frame, self._sheets.get_player_list())
        add_name_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_name)
        add_name_entry.bind("<<ComboboxSelected>>", self.get_player_data)
        add_name_entry.bind("<FocusOut>", self.get_player_data)
        add_name_entry.place(x=170, y=125, width=160)
        add_name_entry.focus_set()

        add_race_label = ttk.Label(self._entry_frame, text='Race', font='Calibri 14')
        add_race_label.place(x=340, y=90, width=130)

        add_race_entry = AutocompleteCombobox(self._entry_frame, self.RACES)
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
            self.set_add_level(self._sheets.get_player_level(self.get_add_name()))
            self.set_add_race(self._sheets.get_player_race(self.get_add_name()))
            self.set_add_class(self._sheets.get_player_class(self.get_add_name()))

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
                last_row = self._ep_tab.get_ep_rows()
                # Create the index
                index = f'A{last_row + 1}'

                # Add a row to the ep sheet and insert
                # the form items
                self._ep_tab.add_ep_row()

            self._ep_tab.add_ep_item(index, entry_list)

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
        for item in self._sheets.get_player_list():
            if self.get_add_name() == item:
                return True

        return False

    def validate_submit(self):
        err_type = ''
        validation_flags = [False, False, False, False, False, False, False, False]

        for item in self.MONTHS:
            if self.get_add_month() == item:
                validation_flags[0] = True
                break

        num_list = [str(x).zfill(2) for x in range(1, 32)]

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

        num_list = [str(x).zfill(1) for x in range(1, 66)]

        for item in num_list:
            if self.get_add_level() == item:
                validation_flags[4] = True
                break

        # Convert 50+ class titles to base classes for validation purposes
        class_entered = self.get_add_class()

        match class_entered:
            case 'Illusionist' | 'Beguiler' | 'Phantasmist' | 'Coercer':
                class_entered = 'Enchanter'
            case 'Elementalist' | 'Conjurer' | 'Arch Mage' | 'Arch Convoker':
                class_entered = 'Magician'
            case 'Heretic' | 'Defiler' | 'Warlock' | 'Arch Lich':
                class_entered = 'Necromancer'
            case 'Channeler' | 'Evoker' | 'Sorcerer' | 'Archanist':
                class_entered = 'Wizard'
            case 'Vicar' | 'Templar' | 'High Priest' | 'Archon':
                class_entered = 'Cleric'
            case 'Wanderer' | 'Preserver' | 'Hierophant' | 'Storm Warden':
                class_entered = 'Druid'
            case 'Mystic' | 'Luminary' | 'Oracle' | 'Prophet':
                class_entered = 'Shaman'
            case 'Minstrel' | 'Troubadour' | 'Virtuoso' | 'Maestro':
                class_entered = 'Bard'
            case 'Disciple' | 'Master' | 'Grandmaster' | 'Transcendent':
                class_entered = 'Monk'
            case 'Pathfinder' | 'Outrider' | 'Warder' | 'Forest Stalker':
                class_entered = 'Ranger'
            case 'Rake' | 'Blackguard' | 'Assassin' | 'Deceiver':
                class_entered = 'Rogue'
            case 'Cavalier' | 'Knight' | 'Crusader' | 'Lord Protector':
                class_entered = 'Paladin'
            case 'Reaver' | 'Revenant' | 'Grave Lord' | 'Dread Lord':
                class_entered = 'Shadow Knight'
            case 'Champion' | 'Myrmidon' | 'Warlord' | 'Overlord':
                class_entered = 'Warrior'
            case 'Primalist' | 'Animist' | 'Savage Lord' | 'Feral Lord':
                class_entered = 'Beastlord'
            case _:
                class_entered = class_entered

        for item in self.CLASSES:
            if class_entered == item:
                validation_flags[5] = True
                break

        for item in self._sheets.get_player_list():
            if self.get_add_name() == item:
                validation_flags[6] = True
                break

        for item in self.RACES:
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
                self._helper.display_error(err_msg)
                return False

        return True
