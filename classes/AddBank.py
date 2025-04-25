import tkinter as tk
import ttkbootstrap as ttk

# from main import config_tab
from classes.AutocompleteCombobox import AutocompleteCombobox
from classes.Helper import Helper

class AddBank(tk.Toplevel):
    # This class creates the toplevel entry form
    # to add records manually to the ep sheet

    def __init__(self, parent, setup, data_list):
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
        self._helper = Helper()
        self._bank_tab = parent
        self._setup = setup

        self._entry_frame = ttk.Frame(self)
        self._add_officer = tk.StringVar()
        self._add_mule = tk.StringVar()
        self._add_item = tk.StringVar()
        self._add_qty = tk.StringVar()
        self._add_notes = tk.StringVar()
        self._add_id = tk.StringVar()

        self.BAD_OFFICER_ERROR = "The officer entered must match the log file selected.\nPlease try again."
        self.BAD_MULE_ERROR = "The mule entered must match a mule on the Configure tab.\nPlease try again."
        self.EMPTY_FIELD_ERROR = "All fields must be filled out.\nPlease try again."

        # This will either be an empty list or it will
        # contain data the user is trying to edit
        self._data_list = data_list

        self.populate_fields(data_list)
        self.create_widgets()

    # Public class member getters
    def get_add_officer(self):
        return self._add_officer.get()

    def get_add_mule(self):
        return self._add_mule.get()

    def get_add_item(self):
        return self._add_item.get()

    def get_add_qty(self):
        return self._add_qty.get()

    def get_add_notes(self):
        return self._add_notes.get()

    def get_add_id(self):
        return self._add_id.get()

    # Public class member setters
    def set_add_officer(self, value):
        self._add_officer.set(value)

    def set_add_mule(self, value):
        self._add_mule.set(value)

    def set_add_item(self, value):
        self._add_item.set(value)

    def set_add_qty(self, value):
        self._add_qty.set(value)

    def set_add_notes(self, value):
        self._add_notes.set(value)

    def set_add_id(self, value):
        self._add_id.set(value)

    def populate_fields(self, data_list):
        # This function auto populates some of the fields
        # on the AddRecord form, to save time
        # Parameters: self (inherit from AddRecord parent)
        # Return: none

        # If data_list contains anything, this is an edit
        # existing record rather than an add new record,
        # so pull in the relevant data fields
        if len(data_list) > 0:
            self.set_add_officer(data_list[0])
            self.set_add_mule(data_list[1])
            self.set_add_item(data_list[2])
            self.set_add_qty(data_list[3])
            self.set_add_notes(data_list[4])
            self.set_add_id(data_list[5])

    def create_widgets(self):
        # This function lays out the widgets on the form
        # Parameters: self (inherit from AddRecord parent)
        # Return: none

        self._entry_frame.pack(expand=True, fill='both')

        add_officer_label = ttk.Label(self._entry_frame, text='Officer', font='Calibri 14')
        add_officer_label.place(x=10, y=10, width=90)

        add_officer_entry = ttk.Entry(self._entry_frame)
        add_officer_entry.configure(font=('Calibri', 12), textvariable=self._add_officer)
        self.set_add_officer(self._setup.get_officer())
        add_officer_entry.place(x=10, y=45, width=100)

        add_mule_label = ttk.Label(self._entry_frame, text='Mule', font='Calibri 14')
        add_mule_label.place(x=120, y=10, width=120)

        add_mule_entry = AutocompleteCombobox(self._entry_frame, self._bank_tab.get_bank_mule_list())
        add_mule_entry.configure(font=('Calibri', 12), height=40, textvariable=self._add_mule)
        add_mule_entry.place(x=120, y=45, width=120)

        add_item_label = ttk.Label(self._entry_frame, text='Spell/Item', font='Calibri 14')
        add_item_label.place(x=250, y=10, width=170)

        add_item_entry = ttk.Entry(self._entry_frame)
        add_item_entry.configure(font='Calibri 12', textvariable=self._add_item)
        add_item_entry.place(x=250, y=45, width=170)

        add_qty_label = ttk.Label(self._entry_frame, text='Qty', font='Calibri 14')
        add_qty_label.place(x=430, y=10, width=40)

        add_qty_entry = ttk.Entry(self._entry_frame)
        add_qty_entry.configure(font='Calibri 12', textvariable=self._add_qty)
        add_qty_entry.place(x=430, y=45, width=40)

        add_notes_label = ttk.Label(self._entry_frame, text='Notes', font='Calibri 14')
        add_notes_label.place(x=10, y=90, width=380)

        add_notes_entry = ttk.Entry(self._entry_frame)
        add_notes_entry.configure(font=('Calibri', 12), textvariable=self._add_notes)
        add_notes_entry.place(x=10, y=125, width=380)

        add_id_label = ttk.Label(self._entry_frame, text='ID', font='Calibri 14')
        add_id_label.place(x=400, y=90, width=70)

        add_id_entry = ttk.Entry(self._entry_frame)
        add_id_entry.configure(font=('Calibri', 12), textvariable=self._add_id)
        add_id_entry.place(x=400, y=125, width=70)

        clear_add_button = ttk.Button(self._entry_frame, text='Clear', style='primary.Outline.TButton',
                                      command=self.clear_form)
        clear_add_button.place(x=175, y=190, width=140)

        submit_button = ttk.Button(self._entry_frame, text='Submit', style='primary.Outline.TButton',
                                   command=self.submit)
        submit_button.place(x=325, y=190, width=140)

    # def get_player_data(self, event):
    #     # This function auto-populates the class and
    #     # races boxes, according to the name entered
    #     # in the name box
    #     # Parameters: self (inherit from AddRecord parent),
    #     #             event (object generated by event)
    #     # Return: none
    #     if self.validate_name():
    #         self.set_add_level(sheets.get_player_level(self.get_add_name()))
    #         self.set_add_race(sheets.get_player_race(self.get_add_name()))
    #         self.set_add_class(sheets.get_player_class(self.get_add_name()))

    def clear_form(self):
        # This function clears out all the
        # fields on the entry form
        # Parameters: self (inherit from AddForm parent)
        # Return: none

        self.set_add_officer('')
        self.set_add_mule('')
        self.set_add_item('')
        self.set_add_qty('')
        self.set_add_notes('')
        self.set_add_id('')

    def submit(self):
        # This function captures the form data
        # into a list and inserts it into the ep sheet
        # Parameters: self (inherit from AddForm parent)
        # Return: none

        if self.validate_submit():
            # Create the list of form fields
            entry_list = [self.get_add_officer(), self.get_add_mule(), self.get_add_item(), self.get_add_qty(),
                          self.get_add_notes(), self.get_add_id()]

            # If data_list is not empty, this is an edit, so get
            # row number from data_list
            if len(self._data_list) > 0:
                index = self._data_list[6]
            else:
                # Otherwise, obtain the last row of the ep sheet
                # so we know where to insert the new data
                last_row = self._bank_tab.get_bank_rows()
                # Create the index
                index = f'A{last_row + 1}'

                # Add a row to the ep sheet and insert
                # the form items
                self._bank_tab.add_bank_row()

            self._bank_tab.add_bank_item(index, entry_list)

            # Close the toplevel form, returning
            # control to the main window
            self.destroy()
            self.update()

    def validate_submit(self):
        if self.get_add_officer() != self._setup.get_officer():
            self._helper.display_error(self.BAD_OFFICER_ERROR)
            return False

        if self.get_add_mule() not in self._bank_tab.get_bank_mule_list():
            self._helper.display_error(self.BAD_MULE_ERROR)
            return False

        if len(self.get_add_item()) < 1:
            self._helper.display_error(self.EMPTY_FIELD_ERROR)
            return False

        if len(self.get_add_qty()) < 1:
            self._helper.display_error(self.EMPTY_FIELD_ERROR)
            return False

        if len(self.get_add_notes()) < 1:
            self._helper.display_error(self.EMPTY_FIELD_ERROR)
            return False

        if len(self.get_add_id()) < 1:
            self._helper.display_error(self.EMPTY_FIELD_ERROR)
            return False

        return True
