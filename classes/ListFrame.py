import tkinter as tk
import ttkbootstrap as ttk
import platform

from classes.GoogleSheets import GoogleSheets
from classes.AutocompleteCombobox import AutocompleteCombobox

class ListFrame(ttk.Frame):
    # A Tkinter scrollable canvas with internal widgets.
    # Created by Christian Koch; original source code at
    # https://github.com/clear-code-projects/tkinter-complete
    # Modified by Timothy Wise for use with Seekers of Souls' EPGP loot
    # system on the Project Quarm emulated Everquest server

    def __init__(self, parent, num_rows):
        super().__init__(master=parent)
        self.pack(expand=True, fill='both')

        self._sheets = GoogleSheets()

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
        player_var = ttk.StringVar(frame, '', player_id)
        gear_var = ttk.StringVar(frame, '', gear_id)
        ratio_var = ttk.StringVar(frame, '', ratio_id)

        # widgets
        count_label = ttk.Entry(frame, font="Calibri 12", foreground="#ffffff")
        count_label.insert(0, f' {index + 1}')
        count_label.configure(state=tk.DISABLED)
        count_label.grid(row=0, column=0, sticky='nsew')

        player_combo = AutocompleteCombobox(frame, self._sheets.get_player_list())
        player_combo.configure(font=('Calibri', 12), height=40, textvariable=player_var)
        player_combo.grid(row=0, column=1, sticky='nsew')
        player_combo.unbind_class('TCombobox', '<MouseWheel>')

        gear_combo = AutocompleteCombobox(frame, self._sheets.get_bid_levels())
        gear_combo.configure(font=('Calibri', 12), height=40, textvariable=gear_var)
        gear_combo.grid(row=0, column=2, sticky='nsew')
        # gear_combo.current(1)

        ratio_label = ttk.Entry(frame, font="Calibri 12", state=tk.DISABLED, foreground="#A9BDBD",
                                textvariable=ratio_var)
        ratio_label.grid(row=0, column=3, sticky='nsew')

        # Insert each value into the 2d list at the
        # unique cell position set up earlier
        self._cells[player_id] = [player_var]
        self._cells[gear_id] = [gear_var]
        self._cells[ratio_id] = [ratio_var]

        return frame
