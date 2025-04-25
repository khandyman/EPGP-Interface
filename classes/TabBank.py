from googleapiclient.errors import HttpError
import tkinter as tk
import ttkbootstrap as ttk
import tksheet
import os.path

# from main import config_tab
# from classes.GoogleSheets import GoogleSheets
from classes.AddBank import AddBank
from classes.Helper import Helper

class TabBank(ttk.Frame):
    # This class creates the Guild Bank Tab, responsible for:
    # - Reading zeal outputfiles for mule characters,
    # - Importing data from the outputfiles
    # - Editing/appending items to the EPGP Log

    def __init__(self, parent, sheets, setup):
        super().__init__(parent)

        # ----- class widget members -----
        self._bank_mule_combo = ttk.Combobox(self)
        self._bank_spells = tk.Radiobutton(self)
        self._bank_items = tk.Radiobutton(self)
        self._bank_sky = tk.Radiobutton(self)
        self._bank_sheet = tksheet.Sheet(self)
        self._bank_clear = ttk.Button(self)
        self._bank_add = ttk.Button(self)
        self._bank_import = ttk.Button(self)
        self._bank_save = ttk.Button(self)

        # ----- class variable members -----
        self._sheets = sheets
        self._helper = Helper()
        self._setup = setup

        self._bank_mule = tk.StringVar()
        self._bank_type = tk.IntVar()
        self._max_row = tk.IntVar()
        self._sky_type = tk.StringVar()
        self._sky_droppables = []
        self._sky_nodrops = []
        self._nodrop_data = []

        # A1 notation ranges for Google Sheets API interactions
        self.ITEM_BANK = "Item Bank!B3:H"
        self.SPELL_BANK = "Spell Bank!B3:G"
        self.SKY_DROPPABLES = "Sky Bank!A3:B"
        self.SKY_NO_DROPS = "Sky Bank!D3:K"

        self.MULE_SELECTION_ERROR = "Please select a mule from the drop down list."
        self.SAVE_BANK_ERROR = "Nothing to save. Please import data first."

        # ----- start-up functions -----
        self.create_widgets()
        self.set_spells_grid()
        self.set_bank_mule_combo()

    # getters
    def get_bank_mule(self):
        return self._bank_mule.get()

    def get_bank_mule_list(self):
        return self._bank_mule_combo['values']

    def get_bank_type(self):
        # _bank_type is a set of three radio buttons
        # sharing the same variable, so we get the
        # variable value, then return a string indicating
        # which radio button is currently selected

        match self._bank_type.get():
            case 1:
                return "Spell Bank"
            case 2:
                return "Item Bank"
            case 3:
                return "Sky Bank"

    def get_max_row(self):
        return self._max_row.get()

    def get_bank_sheet(self):
        return self._bank_sheet.get_sheet_data()

    def get_sky_droppables(self):
        return self._sky_droppables

    def get_sky_nodrops(self):
        return self._sky_nodrops

    def get_nodrop_data(self):
        return self._nodrop_data

    def get_sky_type(self):
        return self._sky_type.get()

    # setters
    def set_bank_mule_combo(self):
        mule_list = []
        # get list of mule file paths from config tab
        file_list = list(self._setup.get_mule_list())
        print(self._setup.get_mule_list())
        for file in file_list:
            # get file name from path
            file_name = os.path.basename(file)
            # slice file name to get character name
            mule_name = file_name[:len(file_name)-14]
            # add character to mule list
            mule_list.append(mule_name)

        # set the combo box to the list of mule names
        self._bank_mule_combo.configure(values=mule_list)
        self.set_bank_mule('')

    def set_bank_mule(self, value):
        self._bank_mule.set(value)

    def set_bank_type(self, value):
        self._bank_type.set(value)

    def set_max_row(self, value):
        self._max_row.set(value)

    def set_bank_sheet(self, sheet_list):
        self._bank_sheet.set_sheet_data(sheet_list)

    def set_sky_droppables(self):
        # if nothing in _sky_droppables list, read
        # in values from EPGP Log; skipped if list
        # contains data to reduce API calls

        if len(self._sky_droppables) < 1:
            droppables = self._sheets.read_values(self.SKY_DROPPABLES, "FORMATTED_VALUE")

            for row in droppables:
                self._sky_droppables.append([item.lower() for item in row])

    def set_sky_nodrops(self):
        # if nothing in _sky_nodrops list, read
        # in values from EPGP Log; skipped if list
        # contains data to reduce API calls

        if len(self.get_sky_nodrops()) < 1:
            no_drops = self._sheets.read_values(self.SKY_NO_DROPS, "FORMATTED_VALUE")

            for row in no_drops:
                self._sky_nodrops.append([row[0].lower(), row[1], row[2], row[3], row[4], row[5], row[6]])

    def set_nodrop_data(self, nodrop_data):
        self._nodrop_data = nodrop_data

    def set_sky_type(self, value):
        self._sky_type.set(value)

    def create_widgets(self):
        # This function creates and places all the widgets
        # on the Guild Bank tab
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        bank_title = ttk.Label(self, text='Guild Bank', font='Calibri 24 bold')
        bank_title.place(x=370, y=0)

        bank_mule_label = ttk.Label(self, text='Mule Selection')
        bank_mule_label.place(x=110, y=50, width=360, height=30)

        bank_type_label = ttk.Label(self, text='Bank Type')
        bank_type_label.place(x=510, y=50, width=170, height=30)

        self._bank_mule_combo.configure(font=('Calibri', 12), height=40, textvariable=self._bank_mule, values=[])
        self._bank_mule_combo.place(x=110, y=80, width=360)

        self._bank_spells.configure(text="Spells", font=('Calibri', 14), variable=self._bank_type, value=1,
                                    background='#073642', foreground='#A9BDBD', command=self.clear_bank)
        self._bank_spells.place(x=510, y=80, width=100)
        self._bank_items.configure(text="Items", font=('Calibri', 14), variable=self._bank_type, value=2,
                                   background='#073642', foreground='#A9BDBD', command=self.clear_bank)
        self._bank_items.place(x=610, y=80, width=100)
        self._bank_sky.configure(text="Sky", font=('Calibri', 14), variable=self._bank_type, value=3,
                                   background='#073642', foreground='#A9BDBD', command=self.clear_bank)
        self._bank_sky.place(x=710, y=80, width=100)
        self._bank_spells.invoke()

        # ttkbootstrap to tksheet mappings: table_bg = inputbg, table_fg = inputfg
        (self._bank_sheet.set_options(table_bg='#073642',
                                        table_fg='#A9BDBD',
                                        header_fg='#ffffff',
                                        index_fg='#ffffff')
         .font(('Calibri', 12, 'normal')))
        (self._bank_sheet.set_options(defalt_row_height=30)
         .height_and_width(width=840, height=368))
        self._bank_sheet.enable_bindings("single_select", "row_select", "right_click_popup_menu", "rc_delete_row")
                                        # "row_select",
                                        # "column_width_resize",
                                        # "arrowkeys",
                                        # "right_click_popup_menu",
                                        # "rc_select",
                                        # "rc_insert_row",
                                        # "rc_delete_row",
                                        # "copy",
                                        # "cut",
                                        # "paste",
                                        # "delete",
                                        # "undo",
                                        # "edit_cell"))
        # Set a double click to activate the edit record function
        self._bank_sheet.set_sheet_data([])
        self._bank_sheet.bind('<Double-Button-1>', self.edit_record)
        self._bank_sheet.place(x=20, y=135)

        self.set_spells_grid()

        ep_sheet_instructions_label = ttk.Label(self,
                                                text='right-click row header to delete; double click anywhere to edit',
                                                font='Calibri 10')
        ep_sheet_instructions_label.place(x=25, y=508)

        self._bank_clear.configure(text='Clear', style='primary.Outline.TButton', command=self.clear_bank)
        self._bank_clear.place(x=270, y=535, width=140)

        self._bank_add.configure(text='Add', style='primary.Outline.TButton', command=self.add_record)
        self._bank_add.place(x=420, y=535, width=140)

        self._bank_import.configure(text='Import', style='primary.Outline.TButton',
                                    command=self.outputfile_import)
        self._bank_import.place(x=570, y=535, width=140)

        self._bank_save.configure(text='Save', style='primary.Outline.TButton',
                                  command=self.insert_bank)
        self._bank_save.place(x=720, y=535, width=140)

    # ----- class functions -----
    def add_record(self):
        # This function opens the Add Record toplevel form
        # so the user can insert rows as needed
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        AddBank(self, self._setup, [])

    def edit_record(self, event):
        # This function grabs data off the sheet in the row
        # clicked by the user, and then opens the add record
        # form so the data can be edited
        # Parameters: self (inherit from TabEP parent
        #             event
        # Return: none

        curr_row = self._bank_sheet.get_currently_selected()[0]
        curr_officer = str(self._bank_sheet.get_cell_data(curr_row, 0))
        curr_mule = str(self._bank_sheet.get_cell_data(curr_row, 1))
        curr_item = str(self._bank_sheet.get_cell_data(curr_row, 2))
        curr_qty = str(self._bank_sheet.get_cell_data(curr_row, 3))
        curr_notes = str(self._bank_sheet.get_cell_data(curr_row, 4))
        curr_id = str(self._bank_sheet.get_cell_data(curr_row, 5))

        data_list = [curr_officer, curr_mule, curr_item, curr_qty, curr_notes, curr_id, curr_row]

        AddBank(self, self._setup, data_list)

    def get_bank_rows(self):
        # This function obtains the current
        # number of rows in the ep sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: get_total_rows (int)

        return self._bank_sheet.get_total_rows()

    def add_bank_row(self):
        # This function adds a row to the
        # bottom of the ep sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self._bank_sheet.insert_row()

    def add_bank_item(self, index, values):
        # This function sets the index row equal
        # to the list of values passed in
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self._bank_sheet.set_data(index, data=[values])

    def clear_bank(self):
        # This function clears all data from the
        # tksheet widget
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        self.set_bank_sheet([])
        self.update_bank_sheet()

    def update_bank_sheet(self):
        # This function determines which radio button
        # is selected and calls the appropriate grid
        # function so that only the relevant columns
        # are displayed
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        match self.get_bank_type():
            case "Spell Bank":
                self.set_spells_grid()
            case "Item Bank":
                self.set_items_grid()
            case "Sky Bank":
                self.set_sky_grid()

    def set_spells_grid(self):
        # This function sets the column widths
        # and row heights for the Spell Bank
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        bank_spells_header = ('Officer', 'Mule', 'Spell', 'Qty', 'Notes', 'Id')
        self._bank_sheet.headers(bank_spells_header)

        (self._bank_sheet
         .column_width(column=0, width=90)      # officer
         .column_width(column=1, width=90)      # mule
         .column_width(column=2, width=170)     # spell
         .column_width(column=3, width=60)      # qty
         .column_width(column=4, width=290)     # notes
         .column_width(column=5, width=90)       # id
         .set_options(default_row_height=30))

    def set_items_grid(self):
        # This function sets the column widths
        # and row heights for the Item Bank
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        bank_items_header = ('Officer', 'Mule', 'Item', 'Qty', 'Notes', 'Id')
        self._bank_sheet.headers(bank_items_header)

        (self._bank_sheet
         .column_width(column=0, width=90)      # officer
         .column_width(column=1, width=90)      # mule
         .column_width(column=2, width=170)     # item
         .column_width(column=3, width=60)      # qty
         .column_width(column=4, width=290)     # notes
         .column_width(column=5, width=90)       # id
         .set_options(default_row_height=30))

    def set_sky_grid(self):
        # This function sets the column widths
        # and row heights for the Sky Bank
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        bank_items_header = ('Officer', 'Mule', 'Item', 'Qty', 'Notes', 'Id')
        self._bank_sheet.headers(bank_items_header)

        (self._bank_sheet
         .column_width(column=0, width=90)      # officer
         .column_width(column=1, width=90)      # mule
         .column_width(column=2, width=170)     # item
         .column_width(column=3, width=60)      # qty
         .column_width(column=4, width=290)     # notes
         .column_width(column=5, width=90)       # id
         .set_options(default_row_height=30))

    def outputfile_import(self):
        # This function is a master function that calls
        # several others in order to populate the sheet
        # with imported data from a mule's outputfile
        # Parameters: self (inherit from TabGP parent)
        # Return: possible Error

        if self.validate_import():
            # build a dictionary of items from
            # an outputfile, filtered to exclude
            # irrelevant lines and items
            bank_dict = self.build_item_dict()
            # convert the item dictionary into a list
            # formatted to match sheet
            item_list = self.build_item_list(bank_dict)

            # add the list items to the sheet
            self.set_bank_sheet(item_list)
            # make sure columns match the type of import
            self.update_bank_sheet()

    def build_item_dict(self):
        # This function reads in outputfile lines
        # and adds data from lines matching filters
        # to master dictionary
        # Parameters: self (inherit from TabGP parent)
        # Return: bank_dict (dictionary)

        # get outputfile file path from config tab
        # list box that matches mule name currently
        # selected in combo box
        inventory_file = self.format_outputfile()
        bank_dict = {}

        # open the file and read it, line by line
        with open(inventory_file) as file:
            line = file.readline()

            # while data in line (indicating EOF not reached)
            while line:
                # determine if current line is a spell
                spell_search = line.find('Spell:')

                # if spell bank selected...
                if self.get_bank_type() == 'Spell Bank':
                    spell_slice = 7

                    # and line contains spell
                    if spell_search != -1:
                        # obtain spell name
                        split = line.split('\t')
                        bank_slot = split[0]
                        spell_name = split[1][spell_slice:]
                        item_id = split[2]

                        # if spell already in dictionary, increment its quantity
                        if spell_name in bank_dict.keys():
                            bank_dict[spell_name]['count'] += 1
                            bank_dict[spell_name]['slots'] = f'{bank_dict[spell_name]['slots']}, {bank_slot}'
                        # otherwise, add spell to dictionary
                        else:
                            bank_dict[spell_name] = {'id': item_id, 'count': int(split[3]), 'slots': bank_slot}
                # if item or sky bank selected
                else:
                    # and line does not contain spell
                    if spell_search == -1:
                        # split line and parse out item name
                        # and first word in line
                        split = line.split('\t')
                        bank_slot = split[0]
                        item_name = self.handle_duplicate(split[1], split[2])
                        item_id = split[2]
                        item_qty = split[3]

                        # if not first line of file, not a currency line, not an empty inventory slot,
                        # and line is not a container
                        if split[0] != 'Location' and split[1] != 'Currency' and split[1] != 'Empty' and int(split[4]) < 1:
                            # if item already in dictionary, increment its quantity
                            if item_name in bank_dict.keys():
                                bank_dict[item_name]['count'] += 1
                                bank_dict[item_name]['slots'] = f'{bank_dict[item_name]['slots']}, {bank_slot}'
                            # otherwise, add item to dictionary
                            else:
                                bank_dict[item_name] = {'id': item_id, 'count': int(item_qty), 'slots': bank_slot}

                line = file.readline()
        return bank_dict

    def build_item_list(self, bank_dict):
        # This function is a master function that calls
        # several others in order to populate the sheet
        # with imported data from a mule's outputfile
        # Parameters: self (inherit from TabGP parent)
        #             bank_dict (dictionary)
        # Return: items_list (list)

        items_list = []

        # loop through dictionary keys
        for key, value in sorted(bank_dict.items()):
            # if this is an item dictionary
            if self.get_bank_type() != 'Spell Bank':
                # run logic to see if item should be added,
                # based on whether item type (sky or not sky)
                # matches radio button selected; if item
                # does not match, skip to next loop index
                if self.add_to_sheet(key) is False:
                    continue

            # format list items to match sheet and add to master list
            item_list = [self._setup.get_officer(), self.get_bank_mule(), key,
                         value['count'], value['slots'], value['id']]
            items_list.append(item_list)

        return items_list

    def format_outputfile(self):
        # This function gets the currently selected
        # mule and pulls the corresponding file path
        # from the config tab list box
        # Parameters: self (inherit from TabGP parent)
        # Return: file (string)

        mule_name = self.get_bank_mule()
        mule_list = self._setup.get_mule_list()

        for file in mule_list:
            # if mule is in current file path return
            # the file path; if not match found
            # function will return empty variable
            if mule_name in file:
                return file

    def add_to_sheet(self, key):
        # This function gets the currently selected
        # mule and pulls the corresponding file path
        # from the config tab list box
        # Parameters: key (string)
        # Return: add_sheet (boolean)

        # obtain sky drop and no drop lists
        self.set_sky_droppables()
        self.set_sky_nodrops()

        # if sky bank selected and item is a sky
        # item, set to true; otherwise, set to false
        if self.get_bank_type() == 'Sky Bank':
            add_sheet = False

            if self.check_sky_lists(key):
                add_sheet = True
        # if item bank selected and item is not a sky
        # item, set to true; otherwise, set to false
        else:
            add_sheet = True

            if self.check_sky_lists(key):
                add_sheet = False

        return add_sheet

    def check_sky_lists(self, key):
        # This function takes in an item name
        # and looks for a matching name in the
        # sky drop and no drop lists
        # Parameters: key (string)
        # Return: found_sky (boolean)

        # set boolean to initial value of false,
        # so that only matches return true
        found_sky = False

        # loop through sky_droppables list
        for row in self.get_sky_droppables():
            # if key found
            if key.lower() in row:
                # flag item as a droppable
                self.set_sky_type('drop')
                found_sky = True
                return found_sky

        # loop through sky_nodrops list
        for row in self.get_sky_nodrops():
            # if key found
            if key.lower() in row:
                # flag item as a no drop
                self.set_sky_type('nodrop')
                self.set_nodrop_data(row)
                found_sky = True
                return found_sky

        # if no match in either list, return false boolean
        return found_sky

    def insert_bank(self):
        # This function creates the dictionary entry
        # of write ranges and write values to send
        # to the Google API for writing to the
        # EPGP Log
        # Parameters: none
        # Return: none

        # if Guild Bank tab passes validation
        if self.validate_save():
            # build list of dictionary entries;
            # each one containing a range and item
            data = self.build_data()

            # set class member max row back to
            # default value of 0; this prevents
            # incorrect max row values if multiple
            # saves are submitted in a single session
            self.set_max_row(0)

            # add data list of dictionaries to master dictionary
            range_body_values = {
                'value_input_option': 'USER_ENTERED',
                'data': data
            }

            # only send API call if data in list
            if len(data) > 0:
                try:
                    # send the HTTP write request to the Sheets API
                    self._sheets.append_values(range_body_values)

                except HttpError as error:
                    print(f"An append error occurred: {error}")
                    self._helper.display_error(f'The Google API has returned this error:\n{error.error_details}')
                    return error

            self.clear_bank()

    def sync_spells(self, log_data, sheet_data):
        # This function compares spell bank tab with app sheet
        # to find spells in log that do not exist on mules;
        # rows found are deleted
        # Parameters: log_data (list)
        #             sheet_data (list)
        # Return: log_data (list)

        # the id of the Spell Bank tab spreadsheet
        sheet_id = '1071255945'
        sheet_mule = sheet_data[0][1]
        # grab the max list index to use in reversing loop
        index = len(log_data) - 1

        # loop through data from log in reverse order, to avoid
        # getting out of sync with ongoing deletions
        for log_row in reversed(log_data):
            # double booleans to signal matching mule but
            # not matching spell
            found_mule = False
            found_spell = False

            # if row is empty, we can't assign values
            if len(log_row) > 0:
                log_mule = log_row[1]

                # compare mule in log row with sheet
                if log_mule == sheet_mule:
                    # flag boolean as true
                    found_mule = True
                    log_spell = log_row[2]

                    # loop through data from sheet
                    for sheet_row in sheet_data:
                        sheet_spell = sheet_row[2]

                        # if spell is found, no deletion necessary,
                        # so flag as true
                        if log_spell == sheet_spell:
                            found_spell = True
                            break

            # if row is empty, because officer forgot to delete it,
            # or if mule from log matches sheet but spell from log not
            # found in sheet
            if len(log_row) < 1 or (found_mule is True and found_spell is False):
                # then build delete JSON payload for sheets API
                request_body = {
                    'requests': [
                        {
                            'deleteDimension': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': index + 2,
                                    'endIndex': index + 3
                                }
                            }
                        }
                    ]
                }

                try:
                    # send a deleteDimension batchUpdate to sheets,
                    # then remove index from log_data to keep it up
                    # to date
                    print(request_body)
                    self._sheets.delete_rows(request_body)
                    log_data.pop(index)
                except HttpError as error:
                    print(f"A delete error occurred: {error}")
                    self._helper.display_error(f'The Google API has returned this error:\n{error.error_details}')

            # decrement index, since this is a reversed loop
            index -= 1

        return log_data

    def build_data(self):
        # This function builds the list of dictionary
        # entries for the save function
        # Parameters: none
        # Return: data (list)

        sheet_data = self.get_bank_sheet()
        # this is a special list for sky items, since
        # drops and no drops are 2 separate columns
        # in the EPGP Log
        log_data_2 = []
        write_range = ''
        write_values = []
        data = []

        # determine radio button selected, then get
        # item data, either from EPGP Log or sky lists
        # also, set quantity column, since it differs
        # from spell/item bank to sky bank
        match self.get_bank_type():
            case "Item Bank":
                log_data = self._sheets.read_values(self.ITEM_BANK, "FORMATTED_VALUE")
                qty_col = 3
            case "Sky Bank":
                log_data = self.get_sky_droppables()
                log_data_2 = self.get_sky_nodrops()
                qty_col = 1
            case _:
                # if spell bank, return result of read_values to syncing algorithm to prune
                # spell bank tab of spells that no longer exist on mules
                log_data = self.sync_spells(self._sheets.read_values(self.SPELL_BANK, "FORMATTED_VALUE"),
                                            sheet_data)
                qty_col = 3

        # loop through each row in the bank sheet
        for sheet_row in sheet_data:
            # obtain variable values for clarity
            officer = sheet_row[0]
            mule = sheet_row[1]
            item = sheet_row[2]
            qty = sheet_row[3]
            notes = sheet_row[4]

            # if current item is a sky no drop item
            if (self.check_sky_lists(item) is True
                    and self.get_sky_type() == 'nodrop'):
                # use the sky no drop list to obtain correct row
                row_count = self.get_log_row([sheet_row], log_data_2, qty_col)
            # otherwise (spells, items, and sky droppables) use the log_data
            # list obtained above
            else:
                row_count = self.get_log_row([sheet_row], log_data, qty_col)

            # get_log_returns an int, if this int is -1 it
            # indicates current item should not be entered,
            # so skip to next sheet row
            if row_count == -1:
                continue

            # depending on radio button selected, set appropriate
            # write range and write values for spells, items,
            # sky droppables, and sky no drops
            match self.get_bank_type():
                case "Spell Bank":
                    write_range = f'{self.get_bank_type()}!B{row_count}:G{row_count}'
                    write_values = [[officer, mule, item, qty, None, notes]]
                case "Item Bank":
                    write_range = f'{self.get_bank_type()}!B{row_count}:H{row_count}'
                    write_values = [[officer, mule, item, qty, None, None, notes]]
                case "Sky Bank":
                    if self.get_sky_type() == 'nodrop':
                        write_range = f'Sky Bank!D{row_count}:K{row_count}'
                        nodrop_data = self.get_nodrop_data()
                        write_values = [[item, qty, nodrop_data[2], nodrop_data[3], nodrop_data[4], nodrop_data[5],
                                         nodrop_data[6], officer]]
                    else:
                        write_range = f'Sky Bank!A{row_count}:B{row_count}'
                        write_values = [[item, qty]]

            # add range and values, as dictionary values, to dat a list
            data.append({
                'majorDimension': 'ROWS',
                'range': write_range,
                'values': write_values
            }, )

        # print list of dictionary entries, for debugging
        for row in data:
            print(row)

        return data

    def get_log_row(self, sheet_row, log_data, qty_col):
        # This function determines the correct row in
        # the EPGP Log to write item data to
        # Parameters: sheet_row (list)
        #             log_data (list)
        #             qty_col (int)
        # Return: row_count (int)

        found = False
        row_count = 3
        # sheet_row parameter is a 2d list of a single index,
        # so specify the 0th index
        officer = sheet_row[0][0]
        mule = sheet_row[0][1]
        item = sheet_row[0][2]
        qty = sheet_row[0][3]

        # if data in the log_data parameter
        if log_data is not None:
            # loop through each index of log data
            for log_row in log_data:
                if len(log_row) > 0:
                    match self.check_sky_lists(item):
                        # if item is a sky item
                        case True:
                            # look for item in current row, if found
                            # set boolean to true
                            if item.lower() in log_row:
                                if self.get_sky_type() == 'drop':
                                    found = True
                                elif officer.lower() in log_row:
                                    found = True
                        # if item is not a sky item
                        case False:
                            # look for mule and item match in current row,
                            # if found set boolean to true
                            if mule == log_row[1] and item == log_row[2]:
                                found = True

                    # if item was found, then check qty
                    # to see if mule inventory matches
                    # qty already in log; if so, flag
                    # return value to skip this item
                    if found is True:
                        if int(qty) == int(log_row[qty_col]):
                            return -1

                        break

                # increment row counter
                row_count += 1

        # if item was not found, it needs added to
        # bottom of tab
        if found is False:
            # if max row has not been updated, set
            # it to current row count
            if self.get_max_row() == 0:
                self.set_max_row(row_count)
            # if max row already updated, increment it
            # and set row count to it
            else:
                self.set_max_row(self.get_max_row() + 1)
                row_count = self.get_max_row()

        return row_count

    @staticmethod
    def handle_duplicate(item_name, item_id):
        match item_id:
            case '20752':
                return f'{item_name} (0.1)'
            case '20778':
                return f'{item_name} (15.0)'
            case _:
                return item_name

    def validate_save(self):
        # This function validates the save function by
        # checking that data is in bank sheet
        # Parameters: none
        # Return: boolean

        if len(self.get_bank_sheet()) > 0:
            return True
        else:
            self._helper.display_error(self.SAVE_BANK_ERROR)
            return False

    def validate_import(self):
        # This function validates the import function by
        # check that a valid mule has been selected
        # Parameters: none
        # Return: boolean

        if self.get_bank_mule() in self.get_bank_mule_list():
            return True
        else:
            self._helper.display_error(self.MULE_SELECTION_ERROR)
            return False