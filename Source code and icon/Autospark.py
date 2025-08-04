import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from gui_components import AutomationApp

def resource_path(relative_path):
    """ Get absolute path to resource (for PyInstaller bundles) """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def main():
    """Main entry point of the application"""
    # Create the root window
    root = tk.Tk()

    # Only create PhotoImage AFTER root is initialized
    icon_path = resource_path("autospark.ico")
    root.iconbitmap(icon_path)

    root.title("AutoSpark")
    root.geometry("900x600")
    root.minsize(800, 500)

    # Apply theme
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
    style.configure("Treeview.Heading", background="#e0e0e0", foreground="black", relief="raised", font=('Arial', 9, 'bold'))
    style.map('Treeview', background=[('selected', '#e0e0ff')], foreground=[('selected', 'black')])
    style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
    style.configure("TButton", font=("Segoe UI", 10), padding=6)

    app = AutomationApp(root)
    root.mainloop()

# Don't put PhotoImage here. Only call main()
if __name__ == "__main__":
    try:
        import logging
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[logging.StreamHandler()])
        logger = logging.getLogger()
        logger.info("Starting AutoSpark app...")

        class Logger:
            def __init__(self, filename="debug_output.txt"):
                self.terminal = sys.__stdout__
                self.log = open(filename, "w")

            def write(self, message):
                if self.terminal:
                    try:
                        self.terminal.write(message)
                    except Exception:
                        pass
                self.log.write(message)
                self.log.flush()

            def flush(self):
                if self.terminal:
                    try:
                        self.terminal.flush()
                    except Exception:
                        pass
                self.log.flush()

        sys.stdout = Logger()
        sys.stderr = Logger("error_output.txt")

        try:
            logger.info("Calling main()")
            main()
        except Exception as e:
            logger.error(f"ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
            raise

    except Exception as e:
        import traceback
        messagebox.showerror("Error", f"An error occurred:\n\n{traceback.format_exc()}")
