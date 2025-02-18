try:
    from MainFrame import *
except ModuleNotFoundError:  # User opted not to install pypdf (and package was not already installed)
    raise SystemExit

import sys
import os


# Global variables
version = 1.0

# Create folder for files if not already present
if not os.path.exists(f"{os.path.dirname(sys.argv[0])}\\Files"):
    os.makedirs(f"{os.path.dirname(sys.argv[0])}\\Files")
init_path = os.path.dirname(sys.argv[0]) + "\\Files"
sys.path.append(init_path)

def get_preferences(preferences: dict) -> str:
    """
    Read in preferences from file.

    :param preferences: Dictionary of program preferences
    :return: String with full pathname of preferences file
    """

    # Check if preferences file exists
    preferences_file = f"{init_path}\\Preferences.txt"
    load_defaults = False
    if not os.path.exists(preferences_file):
        manual_sel = messagebox.askyesno(title="Missing Preferences",
                                         message="The Preferences file could not be located. Would you like to select "
                                                 "the file manually?")
        # User opted not to select file, so load in defaults
        if not manual_sel:
            load_defaults = True

        # User opted to select file
        else:
            is_txt = False
            messagebox.showinfo(title="Select Preferences", message="Select the \".txt\" file containing the "
                                                                    "preferences.")
            while not is_txt:
                preferences_file = filedialog.askopenfilename(initialdir=init_path)

                # No file selected, so confirm retry
                if preferences_file == "":
                    retry = messagebox.askyesno(title="No File Selected", message="No file was selected. Would you "
                                                                                  "like to retry?")
                    if not retry:
                        preferences_file = ""
                        is_txt = True
                        load_defaults = True
                    continue

                # Confirm that file is .txt
                if preferences_file[-4:] != ".txt":
                    messagebox.showerror(title="Invalid Extension", message="The selected file has the wrong extension."
                                                                            " Please select a \".txt\" file.")
                    preferences_file = ""
                    continue

                # File is valid
                is_txt = True

    # Load in from selected file
    if not load_defaults and preferences_file:
        with (open(preferences_file, "r")) as pref:
            for line in iter(pref.readline, ''):
                key, value = line.split(":")
                value = value.strip()

                # Convert values to integers if possible
                try:
                    value = int(value)
                except ValueError:
                    pass

                # Convert Enabled/Disabled to booleans
                value = True if value == "Enabled" else value
                value = False if value == "Disabled" else value

                preferences.update({key: value})
        return preferences_file

    # Load in defaults
    preferences.update({"Font Type": "Times New Roman",
                        "Font Size": 12,
                        "Dark Mode": True,
                        "Combine Non-Sequential File Selections on Move": "Ask",
                        "Add Blank Page Between Files": True})

    return preferences_file

def __main__():
    # Setup main window
    win = Tk()
    win.title(f"PDF Combiner v. {version}")
    win.resizable(False, True)

    # Get preferences
    preferences = {}
    pref_file = get_preferences(preferences)
    dark_mode = preferences["Dark Mode"]

    # Set up styles
    style = ttk.Style()
    if dark_mode:
        win.configure(bg="black")
        style.configure("TLabel", font=(preferences["Font Type"], preferences["Font Size"]), foreground="white",
                        background="black")
        style.configure("TFrame", foreground="white", background="black")
        style.configure("TLabelframe.Label", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="white", background="black")
        style.configure("TLabelframe", background="black")
        style.configure("TRadiobutton", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="white", background="black")
        style.configure("TEntry", font=(preferences["Font Type"], preferences["Font Size"]))
        style.configure("TCheckbutton", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="white", background="black")
        style.configure("TCombobox", font=(preferences["Font Type"], preferences["Font Size"]))

    else:
        style.configure("TLabel", font=(preferences["Font Type"], preferences["Font Size"]), foreground="black",
                        background="SystemButtonFace")
        style.configure("TFrame", foreground="black", background="SystemButtonFace")
        style.configure("TLabelframe.Label", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="black", background="SystemButtonFace")
        style.configure("TLabelframe", background="SystemButtonFace")
        style.configure("TRadiobutton", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="black", background="SystemButtonFace")
        style.configure("TEntry", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="black", background="SystemButtonFace")
        style.configure("TCheckbutton", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="black", background="SystemButtonFace")
        style.configure("TCombobox", font=(preferences["Font Type"], preferences["Font Size"]),
                        foreground="black", background="SystemButtonFace")

    # Populate MainFrame
    MainFrame(win, preferences, pref_file)

    win.mainloop()

if __name__ == __main__():
    __main__()