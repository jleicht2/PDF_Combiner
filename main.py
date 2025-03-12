try:
    from MainFrame import *
except ModuleNotFoundError:  # User opted not to install pypdf (and package was not already installed)
    raise SystemExit

import sys
import os
import win32com.client


# Global variables
version = 1.2

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
    preferences_file = f"{init_path}\\Preferences.pkl"
    load_defaults = False

    # Upgrading from older version: add new keys and remove text file after loading
    if not os.path.exists(preferences_file) and os.path.exists(f"{init_path}\\Preferences.txt"):
        # Load existing preferences
        with (open(f"{init_path}\\Preferences.txt", "r")) as pref:
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

        # Add new keys
        preferences.update({"Shortcut Prompt": True})
        preferences.update({"Desktop Shortcut": ""})
        preferences.update({"Start Menu Shortcut": ""})

        messagebox.showinfo(title="Preferences File Changed", message="The Preferences file has been updated to a new "
                                                                      "format to align with other program changes.")
        os.remove(f"{init_path}\\Preferences.txt")

        # Write output to pickle file
        with (open(preferences_file, "wb")) as pref:
            pickle.dump(preferences, pref)

    elif not os.path.exists(preferences_file):
        manual_sel = messagebox.askyesno(title="Missing Preferences",
                                         message="The Preferences file could not be located. \n"
                                                 "Choose Yes to select a previously-created Preferences file.\n"
                                                 "Choose No to create a new Preferences file with the default options.")
        # User opted not to select file, so load in defaults
        if not manual_sel:
            load_defaults = True

        # User opted to select file
        else:
            is_txt = False
            messagebox.showinfo(title="Select Preferences", message="Select the \".pkl\" file containing the "
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
                if preferences_file[-4:] != ".pkl":
                    messagebox.showerror(title="Invalid Extension", message="The selected file has the wrong extension."
                                                                            " Please select a \".pkl\" file.")
                    preferences_file = ""
                    continue

                # File is valid
                is_txt = True

    # Load in from selected file
    if not load_defaults and preferences_file:
        with (open(preferences_file, "rb")) as pref:
            preferences.update({key: value for key, value in pickle.load(pref).items()})
        return preferences_file

    # Load in defaults
    preferences.update({"Font Type": "Times New Roman",
                        "Font Size": 12,
                        "Dark Mode": True,
                        "Combine Non-Sequential File Selections on Move": "Ask",
                        "Add Blank Page Between Files": True,
                        "Shortcut Prompt": True,
                        "Desktop Shortcut": "",
                        "Start Menu Shortcut": ""})

    return preferences_file


def add_shortcuts(preferences: dict) -> None:
    """
    Check for existence of desktop and start menu shortcuts and create or edit shortcuts as needed.

    :param preferences: Dictionary of program configuration options
    :return:
    """

    # Return if shortcuts should not be created and no shortcuts exist
    if (not preferences["Shortcut Prompt"] and preferences["Desktop Shortcut"] == "" and
            preferences["Start Menu Shortcut"] == ""):
        return

    # Sub-functions to create a shortcut
    def create_shortcut(sc_path: str) -> None:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(sc_path)
        shortcut.Targetpath = sys.argv[0]
        shortcut.save()

    # Check for existence of previously-defined shortcut and recreate if necessary
    if not os.path.exists(preferences["Desktop Shortcut"]) and preferences["Desktop Shortcut"] != "":
        recreate = messagebox.askyesno(title="Missing Desktop Shortcut",
                                       message="The desktop shortcut created previously could not be found. Would you "
                                               "like to create a new one?")
        if recreate:
            path = os.path.join(os.path.expanduser("~"), "Desktop", "PDF Combiner.lnk")
            create_shortcut(path)
            preferences["Desktop Shortcut"] = path
        else:
            preferences["Desktop Shortcut"] = ""

    if not os.path.exists(preferences["Start Menu Shortcut"]) and preferences["Start Menu Shortcut"] != "":
        recreate = messagebox.askyesno(title="Missing Start Menu Shortcut",
                                       message="The Start Menu shortcut created previously could not be found. Would "
                                               "you like to create a new one?")
        if recreate:
            path = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu",
                                "Programs", "PDF Combiner.lnk")
            create_shortcut(path)
            preferences["Start Menu Shortcut"] = path
        else:
            preferences["Start Menu Shortcut"] = ""

    # Ask to create new shortcuts
    if not preferences["Shortcut Prompt"]:
        return

    if preferences["Desktop Shortcut"] == "":
        desktop = messagebox.askyesno(title="Create Desktop Shortcut", message="Should a shortcut be added to the "
                                                                               "desktop?")
        if desktop:
            path = os.path.join(os.path.expanduser("~"), "Desktop", "PDF Combiner.lnk")
            create_shortcut(path)
            preferences["Desktop Shortcut"] = path

    if preferences["Start Menu Shortcut"] == "":
        start = messagebox.askyesno(title="Create Start Menu Shortcut", message="Should a shortcut be added to the "
                                                                                "start menu?")
        if start:
            path = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu",
                                "Programs", "PDF Combiner.lnk")
            create_shortcut(path)
            preferences["Start Menu Shortcut"] = path

    custom = messagebox.askyesno(title="Custom Location", message="Would you like to add a shortcut to a folder of "
                                                                  "your choosing?")
    if custom:
        messagebox.showinfo(title="Select Folder", message="Select the folder in which to place the shortcut.")
        fldr = ""
        retry = True
        while retry:
            fldr = filedialog.askdirectory()
            if fldr == "":
                retry = messagebox.askyesno(title="No Folder Selected", message="No folder was selected. Would you "
                                                                                "like to retry?")
                continue
            retry = False
        if fldr != "":
            fldr = os.path.join(fldr, "PDF Combiner.lnk")
            create_shortcut(fldr)

    preferences["Shortcut Prompt"] = False


def __main__():
    # Setup main window
    win = Tk()
    win.title(f"PDF Combiner v. {version}")
    win.resizable(False, True)

    # Get preferences
    preferences = {}
    pref_file = get_preferences(preferences)
    dark_mode = preferences["Dark Mode"]

    # Set up shortcuts if needed
    add_shortcuts(preferences)

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
