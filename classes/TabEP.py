import os
from file_read_backwards import FileReadBackwards
from googleapiclient.errors import HttpError
from datetime import datetime
import ttkbootstrap.dialogs
import tkinter.messagebox
import tkinter as tk
import ttkbootstrap as ttk
import tksheet
from dotenv import load_dotenv
import discord
import asyncio

from classes.AutocompleteCombobox import AutocompleteCombobox
from classes.AddEP import AddEP
from classes.Helper import Helper

class TabEP(ttk.Frame):
    # This class creates the EP Tab, responsible for:
    # - Retrieving log file time stamps and displaying attendance data
    # - Sending attendance data to the EPGP Log
    # - Manual input is also available

    def __init__(self, parent, sheets, setup):
        super().__init__(parent)
        self.parent = parent

        # ----- private class members -----
        self._helper = Helper()
        self._sheets = sheets
        self._setup = setup

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

        load_dotenv()
        self.TOKEN = os.getenv('DISCORD_TOKEN')
        self._meeting_list = []

        self.configure(width=880, height=630)
        self.place(x=0, y=0)

        # build and configure the widget layout
        self.create_widgets()
        self.set_ep_grid()

        # A1 notation ranges for Google Sheets API interactions
        self.EP_LOG_RANGE = "EP Log!B3:B"

        self.NO_VOICE_ERROR = "No users found in Guild-Meeting channel"
        self.RAID_FOUND_ERROR = 'No raids found in log file. You will not be able to retrieve EP data.'
        self.EMPTY_EP_ERROR = 'Please enter at least one full line of data in the EP sheet and try again.'
        self.TYPE_EP_ERROR = 'Please select a point type and try again.'
        self.SPELLING_EP_ERROR = '] not found in character list. \nAre you sure you want to continue?'
        self.SCAN_ERROR = 'Please select an available time stamp'

        #self.look_for_raids(True)
        self.clear_ep()

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

    def get_meeting_list(self):
        return self._meeting_list

    # ----- public class member setters -----
    def set_ep_time(self, ep_value):
        self._ep_time.set(ep_value)

    def set_ep_combo_values(self, ep_value):
        self._ep_time_combo.configure(values=ep_value)

    def set_ep_type(self, ep_value):
        self._ep_type.set(ep_value)

    def set_ep_sheet(self, ep_value):
        self._ep_sheet.set_sheet_data(ep_value)

    def set_meeting_list(self, meeting_list):
        self._meeting_list = meeting_list

    # ----- widget setup -----
    def create_widgets(self):
        # This function configures and places the tkinter
        # widgets onto the EP Tab frame
        # Parameters:  self (inherit from TabEP parent)
        # Return: none

        ep_title = ttk.Label(self, text='EP Entry', font='Calibri 24 bold')
        ep_title.place(x=380, y=0)

        ep_time_label = ttk.Label(self, text='Import Filter')
        ep_time_label.place(x=110, y=50, width=360, height=30)

        ep_type_label = ttk.Label(self, text='Point Type')
        ep_type_label.place(x=590, y=50, width=170, height=30)

        self._ep_time_combo.configure(font=('Calibri', 12), height=40, textvariable=self._ep_time, values=[])
        self._ep_time_combo.place(x=110, y=80, width=360)

        ep_type_entry = AutocompleteCombobox(self, self._sheets.get_effort_points())
        ep_type_entry.configure(font=('Calibri', 12), height=40, textvariable=self._ep_type)
        ep_type_entry.place(x=590, y=80, width=170)

        ep_header = ('Month', 'Date', 'Time', 'Year', 'Level', 'Class', 'Name', 'Race', '.')
        self._ep_sheet.set_header_data(ep_header)
        self._ep_sheet.set_sheet_data([])
        # ttkbootstrap to tksheet mappings: table_bg = inputbg, table_fg = inputfg
        self._ep_sheet.set_options(
            table_bg='#073642',
            table_fg='#A9BDBD',
            header_fg='#ffffff',
            index_fg='#ffffff'
        )
        self._ep_sheet.font(('Calibri', 12, 'normal'))
        self._ep_sheet.set_options(defalt_row_height=30)
        self._ep_sheet.height_and_width(width=840, height=368)
        self._ep_sheet.enable_bindings(
            "single_select",
            "row_select",
            "right_click_popup_menu",
            "rc_delete_row"
        )

        # Set a double click to activate the edit record function
        self._ep_sheet.bind('<Double-Button-1>', self.edit_record)
        self._ep_sheet.place(x=20, y=135)

        ep_sheet_instructions_label = ttk.Label(self,
                                                text='right-click row header to delete; double click anywhere to edit',
                                                font='Calibri 10')
        ep_sheet_instructions_label.place(x=25, y=508)

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

        AddEP(self, self._sheets, data_list)

    # ----- class functions -----
    def set_ep_grid(self):
        # This function sets the column widths
        # and row heights for the main sheet
        # Parameters: self (inherit from TabEP parent)
        # Return: none

        self._ep_sheet.column_width(column=0, width=60)  # month
        self._ep_sheet.column_width(column=1, width=50)  # date
        self._ep_sheet.column_width(column=2, width=90)  # time
        self._ep_sheet.column_width(column=3, width=60)  # year
        self._ep_sheet.column_width(column=4, width=60)  # level
        self._ep_sheet.column_width(column=5, width=150) # class
        self._ep_sheet.column_width(column=6, width=150) # name
        self._ep_sheet.column_width(column=7, width=100)  # race1
        self._ep_sheet.column_width(column=8, width=40)  # race2
        self._ep_sheet.set_options(default_row_height=30)

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
            point_type = self.get_ep_type()

            # Loop through each line in the guild list,
            # appending the point type each time
            for line in guild_list:
                point_list.append([point_type])

            # Obtain the insertion point for the EP Log;
            # i.e., the first empty cell in column B
            row_count = self._sheets.count_rows(self.EP_LOG_RANGE, 3)

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
                self._sheets.append_values(range_body_values)

            except HttpError as error:
                print(f"An error occurred: {error}")
                self._helper.display_error(f'The Google API has returned this error:\n{error.error_details}')
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

    async def get_discord_users(self):
        client = discord.Client(intents=discord.Intents.all())

        @client.event
        async def on_ready():
            # discord_list = []
            # channel = client.get_all_members()

            # for member in channel:
            #     if member.bot is False:
            #         discord_list.append(member)
            #
            # self.set_meeting_list(discord_list)

            channel = client.get_channel(1155966760919511176)
            self.set_meeting_list(channel.members)

            await client.close()

        await client.login(self.TOKEN)
        await client.connect()

    def read_discord(self):
        meeting_list = []
        guild_list = []
        loop = asyncio.get_event_loop()
        loop.run_until_complete(self.get_discord_users())

        for member in self.get_meeting_list():
            matches = []
            name = member.display_name

            if name.find(" ") != -1:
                matches.append(name.find(" "))

            if name.find("(") != -1:
                matches.append(name.find("("))

            if name.find("/") != -1:
                matches.append(name.find("/"))

            if name.find(",") != -1:
                matches.append(name.find(","))

            if len(matches) > 0:
                match = min(matches)

                if match > -1:
                    name = name[0:match]

            meeting_list.append(name)

        if len(meeting_list) > 0:
            for name in meeting_list:
                if name in self._sheets.get_player_list():
                    time = datetime.now()

                    curr_month = time.strftime('%b')
                    curr_date = time.strftime('%#d')
                    curr_time = time.strftime('%H:%M:%S')
                    curr_year = time.strftime('%Y')
                    curr_level = self._sheets.get_player_level(name)
                    curr_race = self._sheets.get_player_race(name)
                    curr_name = name
                    curr_class = self._sheets.get_player_class(name)

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
            self._helper.display_error(self.NO_VOICE_ERROR)

        return guild_list

    def read_log(self, read_month, read_date, read_time):
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
        with (FileReadBackwards(self._setup.get_log_file(), encoding="utf-8") as frb):
            while True:
                # try block, essentially just to catch and ignore
                # Unicode decode errors; not doing anything with
                # the error, because I don't care about the data;
                # it's just weird characters in the log file
                try:
                    line = frb.readline()

                    if not line:
                        break

                    # If the guild_tag is on the line, plus the month/date/time match what
                    # the user specified, then a match is found """
                    if guild_tag in line:
                        if read_month in line and read_date in line and read_time in line:
                            found_line = True
                            # Remove/replace unwanted formatting from EQ from the line
                            # This is annoying, but the only way I could find to avoid
                            # space delimiting causing problems with two word titles
                            # in the split command below, is to temporarily replace the
                            # spaces with underscores, perform the split, then convert
                            # back to spaces
                            stripped_text = (str(line)
                                             .replace('[', '')
                                             .replace(']', '')
                                             .replace('(', '')
                                             .replace(')', '')
                                             .replace('Shadow Knight ', 'Shadow_Knight ')
                                             .replace('High Priest', 'High_Priest')
                                             .replace('Arch Mage', 'Arch_Mage')
                                             .replace('Arch Convoker', 'Arch_Convoker')
                                             .replace('Grave Lord', 'Grave_Lord')
                                             .replace('Dread Lord', 'Dread_Lord')
                                             .replace('Savage Lord', 'Savage_Lord')
                                             .replace('Feral Lord', 'Feral_Lord')
                                             .replace('Storm Warden', 'Storm_Warden')
                                             .replace('Arch Lich', 'Arch_Lich')
                                             .replace('Lord Protector', 'Lord_Protector')
                                             .replace('Forest Stalker', 'Forest_Stalker')
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
                            list_text = removed_end_text.split()

                            match list_text[5]:
                                case 'Shadow_Knight':
                                    list_text[5] = 'Shadow Knight'
                                case 'Grave_Lord':
                                    list_text[5] = 'Grave Lord'
                                case 'Dread_Lord':
                                    list_text[5] = 'Dread Lord'
                                case 'High_Priest':
                                    list_text[5] = 'High Priest'
                                case 'Arch_Mage':
                                    list_text[5] = 'Arch Mage'
                                case 'Arch_Convoker':
                                    list_text[5] = 'Arch Convoker'
                                case 'Savage_Lord':
                                    list_text[5] = 'Savage Lord'
                                case 'Feral_Lord':
                                    list_text[5] = 'Feral Lord'
                                case 'Storm_Warden':
                                    list_text[5] = 'Storm Warden'
                                case 'Arch_Lich':
                                    list_text[5] = 'Arch Lich'
                                case 'Lord_Protector':
                                    list_text[5] = 'Lord Protector'
                                case 'Forest_Stalker':
                                    list_text[5] = 'Forest Stalker'

                            guild.append(list_text)

                        else:
                            # If found_line is already true but did not find guild tag and time
                            # stamp on the line, then we have reached the end of the /who guild
                            # output, so break out of the loop and return the guild list
                            if found_line:
                                break
                except UnicodeDecodeError:
                    continue

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

        if self._setup.get_log_file():
            available_raids = self.find_times()

            if not available_raids:
                tkinter.messagebox.showinfo('No Raids Found',
                                            self.RAID_FOUND_ERROR)
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

    def find_times(self):
        # This function runs a different algorithm to scan
        # user's log file and find instances of /who guild
        # that match raid parameters (i.e., not a /who all guild,
        # at least 12 players with a guild tag in the same zone)
        # and then save the associated time stamps. Algorithm runs
        # until 10 time stamps have been found or 200k log file lines
        # have been scanned, whichever occurs first
        # Parameters: none
        # Return: found_stamps: list

        player_log = self._setup.get_log_file()
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
                try:
                    line = frb.readline()

                except UnicodeDecodeError:
                    continue

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

        AddEP(self, self._sheets, [])

    def validate_ep_entry(self):
        # This function provides the primary validation for the
        # ep tab
        # Parameters: self (inherit from TabEP parent)
        # Return: boolean

        sheet_values = self.get_ep_sheet()
        missing_names = {}
        missing_flag = False

        # If there is not at least one line of data return false
        if len(sheet_values) == 0:
            self._helper.display_error(self.EMPTY_EP_ERROR)
            return False
        else:
            # If less than 8 items are in the first line return
            # false; this simulates the user clicking
            # 'manual entry' and then entering nothing
            if len(sheet_values[0][:15]) < 7:
                self._helper.display_error(self.EMPTY_EP_ERROR)
                return False
            # If the user has not selected an EP type return false
            elif self._ep_type.get() == '':
                self._helper.display_error(self.TYPE_EP_ERROR)
                return False
            # If all those boxes have been ticked, now we verify
            # that all player names entered match entries on the
            # Totals tab of the EPGP Log; if mismatches are found
            # user will be prompted to confirm entry
            else:
                for row in sheet_values:
                    # Set the name found flag to false for each row
                    found_name = False

                    # Loop through the player list
                    for name in self._sheets.get_player_list():
                        # If a match is found, flag it True then exit
                        # the loop
                        if name == row[6]:  # noqa
                            found_name = True
                            break

                    # format two word races into single string
                    if found_name is False:    # noqa
                        if len(row) > 7:    # noqa
                            if row[7] == 'High' or row[7] == 'Dark' or row[7] == 'Wood':    # noqa
                                player_race = row[7] + ' ' + row[8]     # noqa
                            else:
                                player_race = row[7]    # noqa

                        # add current mismatched entry to dictionary and flag that mismatch was found
                        if row[5] == 'ANON':    # noqa
                            missing_names[row[6]] = {'level': '35', 'class': 'Unknown', 'race': 'Unknown'}  # noqa
                        else:
                            missing_names[row[6]] = {'level': row[4], 'class': row[5], 'race': player_race}  # noqa

                        missing_flag = True

                # If a match was never found, display error and
                # prompt for confirmation
                if missing_flag is True:
                    names_to_strings = ', '.join(missing_names)     # this converts the name list into a string
                    confirm = ttk.dialogs.Messagebox.yesno(
                        f"[ {names_to_strings} {self.SPELLING_EP_ERROR}",
                        "Confirm Name Entry", alert=True ) # noqa

                    # if player clicked yes
                    if confirm == 'Yes':
                        for name in missing_names:
                            # update master player list
                            self._sheets.add_player(name)
                            # update master player dictionary
                            self._sheets.update_dict(name, missing_names[name]['level'],
                                               missing_names[name]['class'], missing_names[name]['race'])

                        return True
                    else:
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
            self._helper.display_error(self.SCAN_ERROR)
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
            self._helper.display_error(self.SCAN_ERROR)
            return False
