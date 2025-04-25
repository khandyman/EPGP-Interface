from googleapiclient.errors import HttpError
import tkinter as tk
import ttkbootstrap as ttk
import ttkbootstrap.dialogs
from datetime import datetime

# from main import sheets
from classes.AutocompleteCombobox import AutocompleteCombobox
from classes.ListFrame import ListFrame
from classes.Helper import Helper

class TabGP(ttk.Frame):
    # This class creates the GP Tab, responsible for:
    # - Mimicking the 'Get Priority' tab of the EPGP log,
    #   consisting of entering character names and loot bid
    #   levels, then writing the data to the Log to determine
    #   the winner
    # - Entering loot items won and writing the data to the GP Log tab

    def __init__(self, parent, sheets):
        super().__init__(parent)

        # ----- class members -----
        self._helper = Helper()
        self._sheets = sheets

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

        # A1 notation ranges for Google Sheets API interactions
        self.GP_LOG_RANGE = "GP Log!C2:C"
        self.PRIORITY_TYPE_RANGE = "EP Log!R3:R"
        self.LOOT_WINNER_RANGE = "Get Priority!G4:I4"
        self.GET_PRIORITY_RANGE = "Get Priority!E3:E"

        self.ENTER_GP_ERROR = 'Please enter a date, character, loot and gear level.'
        self.FIND_WINNER_ERROR = 'Please open bidding and enter at least two characters'
        self.COPY_WINNER_ERROR = 'Please find a winner first'
        self.READ_WINNER_ERROR = 'Problem reading winner from Get Priority tab. Did you enter any alts by mistake?'

        self.configure(width=880, height=595)
        self.place(x=0, y=0)

        self.create_widgets()
        # convert date format to match EPGP Log
        self.set_gp_date(datetime.now().strftime("%#m/%#d/%y"))
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
        self._gp_loot.set(gp_value.strip())

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
        list_frame = ListFrame(priority_frame, self._sheets, self._sheets.get_priority_rows())
        priority_frame.place(x=20, y=80, width=540, height=250)

        return list_frame

    def create_widgets(self):
        # This function creates and places all the widgets
        # on the GP tab
        # Parameters: self (inherit from TabGP parent)
        # Return: none

        priority_title = ttk.Label(self, text='Get Priority', font='Calibri 24 bold')
        priority_title.place(x=370, y=0)

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

        priority_winner_text = ttk.Entry(self, font='Calibri 12', state=tk.DISABLED, foreground="#A9BDBD",
                                         textvariable=self._pr_winner)
        priority_winner_text.place(x=620, y=95, width=150, height=40)

        priority_ratio_text = ttk.Entry(self, font='Calibri 12', state=tk.DISABLED, foreground="#A9BDBD",
                                        textvariable=self._pr_ratio)
        priority_ratio_text.place(x=620, y=170, width=90, height=40)

        priority_level_text = ttk.Entry(self, font='Calibri 12', state=tk.DISABLED, foreground="#A9BDBD",
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
        gp_title.place(x=387, y=410)

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

        gp_character_entry = AutocompleteCombobox(self, self._sheets.get_player_list())
        gp_character_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_name)
        gp_character_entry.place(x=155, y=480, width=160)

        # gp_loot_entry = AutocompleteCombobox(self, sheets.get_loot_names())
        self._gp_loot_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_loot)
        self._gp_loot_entry.place(x=330, y=480, width=315)

        gp_level_entry = AutocompleteCombobox(self, self._sheets.get_gear_points())
        gp_level_entry.configure(font=('Calibri', 12), height=40, textvariable=self._gp_level)
        gp_level_entry.place(x=660, y=480, width=200)

        clear_gp_button = ttk.Button(self, text='Clear', style='primary.Outline.TButton',
                                     command=self.clear_gp)
        clear_gp_button.place(x=570, y=535, width=140)

        save_gp_button = ttk.Button(self, text='Save', style='primary.Outline.TButton',
                                    command=self.insert_gp)
        save_gp_button.place(x=720, y=535, width=140)

        # gp_loot had to be created as a class member so it's list contents could be manipulated externally,
        # this forced it to be created before all other widgets and making the tab order undesirable
        # so, we fix the tab order here
        widgets = [gp_date_entry, gp_character_entry, self._gp_loot_entry, gp_level_entry,
                   clear_gp_button, save_gp_button]
        for w in widgets:
            w.lift()

    # def print_frame(self):
    #     for key, value in self.get_list_frame():
    #         if str(key[1:]) == '0':
    #             print(key, value[0].get())

    # ----- class functions -----
    def grab_date(self, event):
        # This function displays the Date Picker widget
        # to the user in response to a left click on
        # the gp_date_entry widget
        # Parameters: self (inherit from TabGP parent)
        #             event - unused, but is auto-generated
        # Return: none

        cal = ttk.dialogs.DatePickerDialog()
        selected = cal.date_selected.strftime("%m/%#d/%y")
        self.set_gp_date(selected)

    def clear_gp(self):
        # This function clears out the entry widgets
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
            print(str(key)[:1])
            match str(key)[:1]:
                case '1':
                    value[0].set('')
                case '2':
                    # set gear level drop downs to the default
                    # value[0].set(sheets.get_bid_levels()[1])
                    value[0].set('')
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
                row_count = self._sheets.count_rows(self.GP_LOG_RANGE, 2)

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
                    self._sheets.append_values(value_range)

                except HttpError as error:
                    print(f"An error occurred: {error.error_details}")
                    self._helper.display_error(f'The Google API has returned this error:\n{error.error_details}')
                    return error

                # Clear out the GP and Get Priority data
                self.clear_gp()
                self.clear_bidders()
        # The only time an error pop-up is displayed
        # outside a validation function
        else:
            self._helper.display_error(self.ENTER_GP_ERROR)

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
        for item in self._sheets.get_loot_names():
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
                self._sheets.update_loot_names(item_entry)
                self.set_gp_loot_entry(self._sheets.get_loot_names())
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

        self._sheets.clear_values()

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
        # from the middRle column if the user has
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
            self._sheets.append_values(range_body_values)

        except HttpError as error:
            print(f"An error occurred: {error}")
            self._helper.display_error(f'The Google API has returned this error:\n{error.error_details}')
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
            winner = self._sheets.read_values(self.LOOT_WINNER_RANGE, 'FORMATTED_VALUE')
            gl_pr = self._sheets.read_values(self.GET_PRIORITY_RANGE, 'FORMATTED_VALUE')
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
                self._helper.display_error(self.READ_WINNER_ERROR)

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
            self._helper.display_error(self.FIND_WINNER_ERROR)

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
            self._helper.display_error(self.COPY_WINNER_ERROR)

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
                for item in self._sheets.get_player_list():
                    # If a match is found set the flag to true
                    if item == character:
                        char_found = True

            # Run the same logic for the bid levels
            if str(key[:1]) == '2':
                gear_level = value[0].get()

                for item in self._sheets.get_bid_levels():
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
