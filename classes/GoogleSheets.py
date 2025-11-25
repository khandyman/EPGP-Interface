from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import tkinter.messagebox
import os.path

from classes.Helper import Helper

class GoogleSheets:
    def __init__(self):
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

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

            self._helper = Helper()

            # If modifying these scopes, delete the file token.json.

            self.SPREADSHEET_ID = "1pu43LSErcxSaaAkaaTrvMi8GfZYV-dRf1qaWyKiveAA"  # official EPGP log
            self.RAW_LEVELS = "EP Log!F3:F"
            self.RAW_CLASSES = "EP Log!G3:G"
            self.RAW_NAMES = "EP Log!H3:H"
            self.RAW_RACES = "EP Log!I3:I"
            self.BID_LEVELS_RANGE = "Get Priority!H19:H"
            self.EFFORT_POINTS_RANGE = "Point Values!B4:B"
            self.GEAR_POINTS_RANGE = "Point Values!H4:H"
            self.LOOT_NAMES_RANGE = "GP Log!E2:E"
            self.LEVEL_RANGE = "Totals!E4:E"
            self.CHARACTER_GEAR_RANGE = "Get Priority!B3:C"

            self.COUNT_PRIORITY_RANGE = "Get Priority!D3:D"
            self.RAIDER_LIST_RANGE = "Totals!C4:C"
            self.PLAYERS_FOUND_ERROR = 'No players found in EPGP Log. Continuing to load with blank player list.'
            self.LOOT_FOUND_ERROR = 'No loot found in EPGP Log. Continuing to load with blank loot list.'

        except HttpError as error:
            print(f"An error occurred: {error}")
            self._helper.display_error("An HTTP error occurred. You may need to request access to EPGP-Interface."
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
            len(self.read_values(self.COUNT_PRIORITY_RANGE, 'FORMULA')))

    # All the rest of the setters just use FORMATTED_VALUE,
    # which just means read the contents of the cell rather
    # than the formula; we are also using the sum function
    # to flatten the 2D lists retrieved into 1D lists
    def set_player_list(self):
        raw_players = self.read_values(self.RAIDER_LIST_RANGE, 'FORMATTED_VALUE')

        if raw_players:
            self._player_list = sum(raw_players, [])
            self._player_dict = self.build_player_dict()
        else:
            tkinter.messagebox.showinfo('No Players Found', self.PLAYERS_FOUND_ERROR)

    def add_player(self, name):
        self._player_list.append(name)

    def update_dict(self, player_name, player_level, player_class, player_race):
        self._player_dict[player_name] = {'level': player_level, 'class': player_class, 'race': player_race}

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
        raw_levels = self.read_values(self.RAW_LEVELS, "FORMATTED_VALUE")
        raw_classes = self.read_values(self.RAW_CLASSES, "FORMATTED_VALUE")
        raw_names = self.read_values(self.RAW_NAMES, 'FORMATTED_VALUE')
        raw_races = self.read_values(self.RAW_RACES, 'FORMATTED_VALUE')

        # Outer loop through Totals tab player list
        for player in self.get_player_list():
            # reset variables to defaults
            player_level = '35'
            player_class = 'Unknown'
            player_race = 'Unknown'

            # Start counter with max of list from column I, need to
            # use this column because it could be shortest
            counter = len(raw_races) - 1

            # Inner loop, decrementing through player names to find a match,
            # but using race list to avoid potential out of range errors
            for raw_race in reversed(raw_races):
                formatted_player = self._helper.strip_garbage(raw_names[counter])

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
                        player_level = (self._helper.strip_garbage(raw_levels[counter]))

                        if player_level == 'ANONYMOUS':
                            player_level = self.get_max_level()

                        player_race = self._helper.strip_garbage(raw_race)

                        if (player_race == 'High'
                                or player_race == 'Dark'
                                or player_race == 'Wood'
                                or player_race == 'Half'):
                            player_race = player_race + ' Elf'

                        player_class = self._helper.strip_garbage(raw_classes[counter])

                        # match found, so exit loop
                        break

                # Decrement the counter
                counter -= 1

            # add to master dictionary with current values, either defaults if
            # no player match found, or values found at 'counter'
            master_dict[player] = {'level': player_level, 'class': player_class, 'race': player_race}

        return master_dict

    def set_effort_points(self):
        self._effort_points = (
            sum(self.read_values(self.EFFORT_POINTS_RANGE, 'FORMATTED_VALUE'), []))

    def set_gear_points(self):
        self._gear_points = (
            sum(self.read_values(self.GEAR_POINTS_RANGE, 'FORMATTED_VALUE'), []))

    def set_bid_levels(self):
        self._bid_levels = (
            sum(self.read_values(self.BID_LEVELS_RANGE, 'FORMATTED_VALUE'), []))

    def set_max_level(self):
        # This function finds the maximum level listed on
        # the Totals tab of the EPGP Log, for autopopulating
        # the AddForm level field
        # Parameters: self (inherit from Sheets parent)
        # Return: curr_max (str)

        curr_max = 0

        # Use the sum function to flatten 2D list into 1D list
        raw_levels = self.read_values(self.LEVEL_RANGE, 'FORMATTED_VALUE')

        if raw_levels:
            level_list = sum(raw_levels, [])

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
        try:
            raw_loot = self.read_values(self.LOOT_NAMES_RANGE, 'FORMATTED_VALUE')

            if raw_loot:
                self._loot_names = (
                    list(set(sum(raw_loot, []))))
                self._loot_names.sort()
            else:
                tkinter.messagebox.showinfo('No Loot Found', self.LOOT_FOUND_ERROR)
        except Exception as error:
            tkinter.messagebox.showinfo('Error', str(error))

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
            creds = Credentials.from_authorized_user_file("token.json", self.SCOPES)
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
                    "credentials.json", self.SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build('sheets', 'v4', credentials=creds)

    def count_rows(self, row_range, offset):
        result = (
            self._service.spreadsheets().values()
            .get(spreadsheetId=self.SPREADSHEET_ID, range=row_range, valueRenderOption='FORMATTED_VALUE')
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
                spreadsheetId=self.SPREADSHEET_ID,
                body=range_body_values)
            .execute()
        )

        # sheets.get_service().spreadsheets().batchUpdate(
        #     spreadsheetId=spreadsheet_id,
        #     body=request_body
        # ).execute()

        # print(f"{(result.get('updates').get('updatedCells'))} cells appended.")

        return result

    def delete_rows(self, request_body):
        # This function deletes a row from the google sheet
        # Parameters: self (inherits from GoogleSheets parent)
        #             request_body (JSON) - deleteDimension payload
        # Return: result object indicating success or failure

        result = (
            self._service.spreadsheets()
            .batchUpdate(
                spreadsheetId=self.SPREADSHEET_ID,
                body=request_body
            ).execute()
        )

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
                spreadsheetId=self.SPREADSHEET_ID,
                range=self.CHARACTER_GEAR_RANGE,
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

        try:
            result = (
                self._service.spreadsheets().values()
                .get(spreadsheetId=self.SPREADSHEET_ID, range=read_range, valueRenderOption=value_render)
                .execute()
            )
            values = result.get("values", [])

            if not values:
                print("No data found.")
                return

            return values
        except HttpError as error:
            print(f"An error occurred: {error}")
            self._helper.display_error("An HTTP error occurred. Please try again in a few moments.")

            return []
