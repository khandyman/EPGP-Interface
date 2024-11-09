Seekers of Souls
EPGP-Interface

Readme / How-to Guide
This document is intended to explain the use of the EPGP-Interface application.  This program was created for use by the leadership of the Seekers of Souls (SoS) guild on the Project Quarm emulated EverQuest server.  It is a front-end interface for the back end EPGP Log that SoS uses to store attendance records and points earned as well as distribute loot obtained during raids and guild events.  The EPGP-Interface was designed specifically to automate as much of the process of taking attendance, determining loot winners, and entering loot awarded as possible.  Where the process could not be automated, an attempt was made to streamline it and make it easier.

Features
•	Integrated Google Oauth security protocols
•	Auto-updating lists pulled from the EPGP Log in real time, including: character names, effort point categories, gear point types, bid levels, and item names
•	Automatic EQ log file scanning to find attendance time stamps
•	One click importing of attendance data from log file
•	Built-in error correction to ensure character names are accurate
•	Easy manual attendance additions with auto-completing lists pulled from the EPGP Log
•	Automated insertion of EP data, GP data, and Get Priority winner determinations
•	Painless GP entry with date formatting and loot name error correction


Getting started
When the EPGP-Interface.zip file is decompressed the following files will be available (credentials.json, EPGP-Interface.exe, Readme.pdf, WindowsIcon.ico).
 
When the program is run for the first time, two more files will be generated (config and token.json).  All of these files must remain in the same directory for the EPGP-Interface to function.  
Token.json is built by the Google Sheets API and will require the user to select a Google account to use with the app.  This is done automatically in a web browser.
 
The EPGP-Interface cannot be run without a log file.  Select the log for the character who will be taking attendance during a SoS event.  After choosing a file the main app screen will be displayed…

The EPGP-Interface consists of three tabs: Effort Points, Gear Points, and Configure.  The Effort Points tab will be used for taking attendance.  The Gear Points tab allows for finding loot winners and entering said loot into the Log.  And the Configure tab is only used (currently) to change the EQ log file.


Entering EP
When entering EP, the first thing to do is select a time stamp from what’s available in the ‘Raid Time Stamps’ drown down box.  After selecting the correct stamp, click ‘Scan Log File.’  This will look for the chosen date/time and compile a list of the SoS members present at that time stamp.  After a few seconds, something like the following will be displayed on the screen…

The data should be checked for accuracy at this point.  Make sure the number of rows matches the number of people in the raid at the time attendance was taken.  If the numbers do not match, check for anyone who was “anonymous” instead of “roleplay” or if people were outside the raid zone when attendance was taken.  These can be added manually.

When the data is confirmed to be accurate and all manual records needed are entered, select the correct “point type” (e.g., Raid – Start, Event Attend, etc.) from the drop down in the upper right corner of the window.  Finally, click “Save EP” at the bottom right.  The spread sheet will disappear, all drop downs will be cleared, and the EP data will be written to the EPGP Log.  The above time stamp, with New Target selected as the point type, will show up in the log like the following picture…

After taking a new attendance in game (e.g., “Raid – Mid” and “Raid – End”), click “Refresh Raids” to update the time stamps available for selection.  Then enter EP as before.


Getting Priority (PR)
To find out who has won a piece of loot, the following process should be followed.  Enter the names of each bidder in the Character drop-down boxes in the Get Priority section.  These boxes contain a list of every SoS raider, and the boxes feature auto-complete, so getting the names in quickly and easily should be a snap.  Then, make sure the gear level each character has bid for is selected.  Finally, click the “Find Winner” button.  The winning bidder will be pulled from the EPGP Log and displayed in the “Winner” section, as shown below…
 
The information in EPGP-Interface is, as with everything else in the application, pulled directly from the EPGP Log.  For example, after clicking the “Find Winner” button as shown above, this is what the Get Priority tab will look like:
 
Getting priority can be done multiple times.  So, if a last-minute bid comes in and the winner has already been determined, just fill add in the new character’s name and click “Find Winner” again.


Entering Gear Points (GP)
After determining the winner for a piece of loot, the GP needs to be entered into the Log.  That is done as follows.  First, click “Copy Winner.”  The winner’s name and gear level will be copied down to the GP Entry section.  If the date needs 

When ready, click the “Save GP” button and the loot will be appended to the bottom of the GP Log tab in the EPGP Log…
 
After entering GP, all fields in the “Get Priority” and “GP Entry” sections will be cleared to make room for new entries.


Configuration
The final tab of the EPGP-Interface application is the Configure tab.  
 
The only function currently in this tab is the “Change Log” function.  This should rarely be done in practice.  One reason would be if an officer changes main characters.  Regardless, if the log file does need to be changed, simply click “Change Log”, and select the new EQ log file.

Appendix A: Theorycrafting for Nerds

