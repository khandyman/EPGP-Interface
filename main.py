from classes.GoogleSheets import GoogleSheets
from classes.Load import Load
from classes.Notebook import Notebook
from classes.TabEP import TabEP
from classes.TabGP import TabGP
from classes.TabBank import TabBank
from classes.TabConfig import TabConfig
from classes.Setup import Setup

# This program was created for use by the leadership of the Seekers of Souls (SoS) guild on the
# Project Quarm emulated EverQuest server.  It is a front-end interface for the back end EPGP Log
# that SoS uses to store attendance records and points earned as well as distribute loot obtained
# during raids and guild events.  The EPGP-Interface was designed specifically to automate as much
# of the process of taking attendance, determining loot winners, and entering loot awarded as possible.
# Where the process could not be automated an attempt was made to streamline it and make it easier.

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
    ep_tab = TabEP(app, sheets, setup)
    gp_tab = TabGP(app, sheets)
    config_tab = TabConfig(app)
    bank_tab = TabBank(app, sheets, setup)

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
    global ep_tab, gp_tab, bank_tab, config_tab

    # global variable to slow down autocomplete
    key_strokes = 0

    # Create the root window
    app = Notebook()

    # Set up the Google Sheets API communication framework
    sheets = GoogleSheets()
    setup = Setup()

    # Create the main tab classes
    # ep_tab = TabEP(app, sheets, setup)
    # gp_tab = TabGP(app, sheets)
    # bank_tab = TabBank(app, setup)
    # config_tab = TabConfig(app)


    # Create loading screen and call it after a brief
    # pause to allow it to display, then run start-up
    load_screen = Load(app)
    load_screen.loading_window.after(100, start_up)

    app.run()
