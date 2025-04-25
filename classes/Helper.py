import ttkbootstrap as ttk
import ttkbootstrap.dialogs

class Helper:
    @staticmethod
    def strip_garbage(value):
        value = (str(value)
                 .replace('[', '')
                 .replace('\'', '')
                 .replace(']', ''))

        return value

    @staticmethod
    def display_error(message):
        # This function serves as a template for all error
        # messages in the application
        ttk.dialogs.Messagebox.show_error(message, 'App Error')