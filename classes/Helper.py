import ttkbootstrap as ttk
import ttkbootstrap.dialogs

class Helper:
    def __init__(self):
        pass

    def strip_garbage(self, value):
        value = (str(value)
                 .replace('[', '')
                 .replace('\'', '')
                 .replace(']', ''))

        return value

    def display_error(self, message):
        # This function serves as a template for all error
        # messages in the application
        ttk.dialogs.Messagebox.show_error(message, 'App Error')