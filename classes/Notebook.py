import ttkbootstrap as ttk

class Notebook(ttk.Window):
    # root window class; parent of all tab classes

    def __init__(self):
        # ----- window setup -----
        super().__init__(themename='solar')
        self.withdraw()
        self.title('EPGP Interface v2.6')
        self.geometry('880x630')
        self.resizable(False, False)
        # main window icon
        self.iconbitmap('WindowsIcon.ico')
        # message box icon
        self.iconbitmap(default='WindowsIcon.ico')

        # Unused - print all colors in theme
        # Colors = self.style.colors
        # for color_label in Colors.label_iter():
        #     color = Colors.get(color_label)
        #     print(color_label, color)

        # todo: revisit ttk style to set style for all widgets at once
        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Calibri', 12))
        style.configure('TCalendar', font=('Calibri', 12))
        style.configure('TLabel', font=('Calibri', 14))
        style.configure('TEntry', font=('Calibri', 12))
        style.configure('primary.Outline.TButton', font=("Calibri", 12))

        # ----- tab setup -----
        # main notebook widget, parent of all frame (tab) widgets
        self.epgp_notebook = ttk.Notebook(self)

    def add_tab(self, class_tab, title):
        # This function creates tabs for the main notebook
        # Parameters: self (inherit from Notebook parent)
        #             class_tab - the name of a class instance
        #             title - the text to display on the window tab
        # Return: none

        self.epgp_notebook.add(class_tab, text=title)
        self.epgp_notebook.place(x=0, y=0, width=880, height=630)

    def run(self):
        # This function executes the window main loop

        self.mainloop()