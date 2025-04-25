import tkinter as tk
import ttkbootstrap as ttk

class AutocompleteCombobox(ttk.Combobox):
    """
    A Tkinter widget that features autocompletion.
    Created by Mitja Martini on 2008-11-29.
    Updated by Russell Adams, 2011/01/24 to support Python 3 and Combobox.
    Updated by Dominic Kexel to use Tkinter and ttk instead of tkinter and tkinter.ttk
    Updated by Timothy Wise to follow Python class conventions for use with
       Seekers of Souls' EPGP loot system on the Project Quarm emulated Everquest server
       Licensed same as original (not specified?), or public domain, whichever is less restrictive.
    """

    def __init__(self, parent, completion_list):
        super().__init__(master=parent)

        self._completion_list = completion_list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.configure(font=('Calibri', 12), height=40)

        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Setup our popup menu
        self.bind('<FocusOut>', self.handle_leave)

    def autocomplete(self, event, delta=0):
        # autocomplete the Combobox, delta may be 0/1/-1 to cycle through possible hits
        if delta:  # need to delete selection otherwise we would fix the current position
            self.delete(self.position, tk.END)
        else:  # set position to end so selection starts where textentry ended
            self.position = len(self.get())

        # update values list with current matches
        _hits = self.update_list()

        # if we have a new hit list, keep this in mind
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        # only allow cycling if we are in a known hit list
        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)

        # now finally perform the auto-completion,
        # if only a single match remains in list
        if self._hits and len(_hits) == 1:
            # get the current cursor position
            curr_index = self.index(tk.INSERT)

            # delete the field contents
            self.delete(0, tk.END)
            # insert the matching list item
            self.insert(0, self._hits[self._hit_index])

            # select the range from the cursor to the end
            # so that any further typing will only
            # over-write what has been auto-populated
            self.select_range(curr_index, tk.END)

            # set cursor position back to where it was before
            # autocomplete, this is key to allow user
            # to keep typing without adding characters
            self.icursor(curr_index)

    def handle_keyrelease(self, event):
        # event handler for the keyrelease event on this widget
        if event.keysym == "BackSpace":
            self.update_list()
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keysym == "Left":
            if self.position < self.index(tk.END):  # delete the selection
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1  # delete one character
                self.delete(self.position, tk.END)
        if event.keysym == "Right":
            self.position = self.index(tk.END)  # go to end (no selection)
        if len(event.keysym) == 1:
            self.autocomplete(event)

    def handle_leave(self, event):
        # clear the current selection; mostly a visual thing
        self.select_clear()

    def update_list(self):
        # start with empty list
        _hits = []

        # if nothing in box, reset list back to default
        if len(self.get()) == 0:
            _hits = self._completion_list
        else:
            # loop through default list
            for element in self._completion_list:
                # look for matches with current entry
                if element.lower().startswith(self.get().lower()):  # Match case insensitively
                    match = False
                    # loop through items in current list
                    for item in _hits:
                        # only add to list if not already there
                        if element.lower() == item.lower():
                            match = True
                            break

                    if match is False:
                        _hits.append(element)
        # set values list to current hit list
        self.configure(values=_hits)

        return _hits
