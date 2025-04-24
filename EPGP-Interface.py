# import os.path
# import platform
# import tkinter.messagebox

# import requests
# import re
# import sys
# import tkinter as tk
# import ttkbootstrap as ttk
# import time
# import msvcrt
# from datetime import datetime
# from tkinter import *
# from tkinter import filedialog
# import tksheet

 # import ttkbootstrap.dialogs
# from file_read_backwards import FileReadBackwards
# from google.auth.transport.requests import Request
# from google.auth.exceptions import RefreshError
# from google.oauth2.credentials import Credentials
# from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError
from classes.GoogleSheets import GoogleSheets
from classes.Load import Load
from classes.Notebook import Notebook
from classes.TabEP import TabEP
from classes.TabGP import TabGP
from classes.TabBank import TabBank
from classes.TabConfig import TabConfig

# This program was created for use by the leadership of the Seekers of Souls (SoS) guild on the
# Project Quarm emulated EverQuest server.  It is a front-end interface for the back end EPGP Log
# that SoS uses to store attendance records and points earned as well as distribute loot obtained
# during raids and guild events.  The EPGP-Interface was designed specifically to automate as much
# of the process of taking attendance, determining loot winners, and entering loot awarded as possible.
# Where the process could not be automated an attempt was made to streamline it and make it easier.


# ---------- constants ----------





# A1 notation ranges for Google Sheets API interactions
GP_LOG_RANGE = "GP Log!C2:C"
EP_LOG_RANGE = "EP Log!B3:B"
PRIORITY_TYPE_RANGE = "EP Log!R3:R"

LOOT_WINNER_RANGE = "Get Priority!G4:I4"
GET_PRIORITY_RANGE = "Get Priority!E3:E"




ITEM_BANK = "Item Bank!B3:H"
SPELL_BANK = "Spell Bank!B3:G"
SKY_DROPPABLES = "Sky Bank!A3:B"
SKY_NO_DROPS = "Sky Bank!D3:K"

# hard-coded lists
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
CLASSES = ['Bard', 'Beastlord', 'Cleric', 'Druid', 'Enchanter', 'Magician', 'Monk', 'Necromancer',
           'Paladin', 'Ranger', 'Rogue', 'Shadow Knight', 'Shaman', 'Warrior', 'Wizard', 'Unknown']
RACES = ['Barbarian', 'Dark Elf', 'Dwarf', 'Erudite', 'Froglok', 'Gnome', 'Half Elf', 'Halfling', 'High Elf',
         'Human', 'Iksar', 'Ogre', 'Troll', 'Vah Shir', 'Wood Elf', 'Unknown']

# validation messages
SCAN_ERROR = 'Please select an available time stamp'
EMPTY_EP_ERROR = 'Please enter at least one full line of data in the EP sheet and try again.'
SPELLING_EP_ERROR = '] not found in character list. \nAre you sure you want to continue?'
TYPE_EP_ERROR = 'Please select a point type and try again.'
ENTER_GP_ERROR = 'Please enter a date, character, loot and gear level.'
FIND_WINNER_ERROR = 'Please open bidding and enter at least two characters'
COPY_WINNER_ERROR = 'Please find a winner first'
READ_WINNER_ERROR = 'Problem reading winner from Get Priority tab. Did you enter any alts by mistake?'


RAID_FOUND_ERROR = 'No raids found in log file. You will not be able to retrieve EP data.'


NO_VOICE_ERROR = "No users found in Guild-Meeting channel"


MULE_SELECTION_ERROR = "Please select a mule from the drop down list."
SAVE_BANK_ERROR = "Nothing to save. Please import data first."
BAD_OFFICER_ERROR = "The officer entered must match the log file selected.\nPlease try again."
BAD_MULE_ERROR = "The mule entered must match a mule on the Configure tab.\nPlease try again."
EMPTY_FIELD_ERROR = "All fields must be filled out.\nPlease try again."

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
    global app, ep_tab, gp_tab, bank_tab, config_tab

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
    config_tab = TabConfig(app, ep_tab, bank_tab)
    bank_tab = TabBank(app)

    # Run an initial scan of the log file to
    # set the available raid time stamps
    load_screen.update_label(' Scanning for raid time stamps')
    ep_tab.look_for_raids(True)
    ep_tab.clear_ep()
    load_screen.update_progress(12.5)

    # Add the tab classes to the main window
    app.add_tab(ep_tab, 'Effort Points')
    app.add_tab(gp_tab, 'Gear Points')
    app.add_tab(bank_tab, 'Guild Bank')
    app.add_tab(config_tab, 'Configure')

    # Restore the main window (it was hidden immediately
    # on creation so the load window could be
    # displayed alone
    app.overrideredirect(False)
    app.attributes('-alpha', 100)
    app.deiconify()

    # Kill the load window, as its no longer needed
    load_screen.loading_window.destroy()


# ---------- main execution code ----------
if __name__ == "__main__":
    # global namespace for main window tabs
    global app, ep_tab, gp_tab, bank_tab, config_tab

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
