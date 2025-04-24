import tkinter as tk
import ttkbootstrap as ttk

class Load:
    def __init__(self, parent):
        # This class creates the load screen window.

        self.loading_window = tk.Toplevel(parent)
        self.loading_window.title("Loading...")
        # Hide the Windows title bar
        self.loading_window.overrideredirect(True)

        parent_x_pos = parent.winfo_rootx()
        parent_y_pos = parent.winfo_rooty()
        geometry = '350x150+%d+%d' % (parent_x_pos + 190, parent_y_pos + 180)
        self.loading_window.geometry(geometry)

        self.label_frame = ttk.Frame(self.loading_window)
        self.label_frame.place(x=0, y=30, width=350, height=35)

        self.progress_label = ttk.Label(self.label_frame, text='Loading...', font='Calibri 14')
        self.progress_label.pack()

        self.progress_frame = ttk.Frame(self.loading_window)
        self.progress_frame.place(x=30, y=100, width=285, height=35)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame,
                                            length=250,
                                            variable=self.progress_var,
                                            maximum=100,
                                            style='TProgressbar')
        self.progress_bar.pack(fill='both', expand=True)

    def update_label(self, value):
        # This function updates the text in the label
        self.progress_label.configure(text='')
        self.progress_label.configure(text=value)
        self.loading_window.update_idletasks()

    def update_progress(self, value):
        # This function updates the value of the progress bar
        new_value = self.progress_var.get() + value
        self.progress_var.set(new_value)
        self.loading_window.update_idletasks()
