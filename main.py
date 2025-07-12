try:
    from MainFrame import *  # Includes Edit Preferences Frame
except ModuleNotFoundError:  # User opted not to install pypdf (and package was not already installed)
    raise SystemExit

import sys
import os
import platform
import shutil

# Global variables
version = "1.4"

# Default preferences
default_pref = {"Font Type": "Times New Roman",
                "Font Size": 12,
                "Dark Mode": True,
                "Combine Non-Sequential File Selections on Move": "Ask",
                "Compress Output": True,
                "Launch File Dialog to Script Folder": True,
                "Add Blank Page Between Files": True,
                "Shortcut Prompt": True,
                "Desktop Shortcut": "",
                "Start Menu Shortcut": ""}

# Create folder for files if not already present
if not os.path.exists(f"{os.path.dirname(sys.argv[0])}\\Files"):
    os.makedirs(f"{os.path.dirname(sys.argv[0])}\\Files")
init_path = os.path.dirname(sys.argv[0]) + "\\Files"
sys.path.append(init_path)


def set_styles(preferences: dict, win: Tk) -> None:
    """
    Set up ttk Styles for win based on user preferences.

    :param preferences: Dictionary of program preferences
    :param win: Tk object in which to set up the styles
    :return:
    """

    style = ttk.Style(win)
    if preferences["Dark Mode"]:
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


def get_preferences(preferences: dict, win: Tk) -> str:
    """
    Read in preferences from file.

    If the file cannot be found, prompt for selecting a previously created preferences file or to create a new one.
    A new file can be created using Edit Preferences Frame or the default_pref dictionary.

    :param preferences: Dictionary of program preferences
    :param win: Parent window for Edit Preferences frame
    :return: String with full pathname of preferences file
    """

    # Check if preferences file exists
    preferences_file = f"{init_path}\\Preferences.pkl"
    load_option = "normal"

    # Upgrading from older version: add new keys and remove text file after loading
    if not os.path.exists(preferences_file) and os.path.exists(f"{init_path}\\Preferences.txt"):
        # Load existing preferences
        load_option = "normal"
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
        for key, value in default_pref.items():
            if key not in preferences.keys():
                preferences.update({key: value})

        messagebox.showinfo(title="Preferences File Changed", message="The Preferences file has been updated to a new "
                                                                      "format to align with other program changes.")

        # Write output to pickle file
        with (open(preferences_file, "wb")) as pref:
            pickle.dump(preferences, pref)

    # File cannot be located
    elif not os.path.exists(preferences_file):
        # Ask if file is in wrong location
        prev_pref = messagebox.askyesno(title="Missing Preferences",
                                         message="The Preferences file could not be located.\n"
                                                 "Have you previously created a preferences file for this app?")

        # No file previously created
        if not prev_pref:
            use_default = messagebox.askyesno(title="Use Default?", message="Should the default preferences be "
                                                                            "used?")
            load_option = "default" if use_default else "user-select"

        # File previously created
        else:
            sel_exist = messagebox.askyesno(title="Select Existing?", message="Would you like to select the existing "
                                                                              "file?")

            # Select existing file
            if sel_exist:
                is_pkl = False
                messagebox.showinfo(title="Select Preferences", message="Select the \".pkl\" file containing the "
                                                                        "preferences.")
                while not is_pkl:
                    preferences_file = filedialog.askopenfilename(initialdir=init_path,
                                                                  filetypes=[("Pickle Files", ".pkl")])

                    # No file selected, so confirm retry
                    if preferences_file == "":
                        retry = messagebox.askyesno(title="No File Selected", message="No file was selected. Would you "
                                                                                      "like to retry?")
                        if not retry:
                            preferences_file = ""
                            is_pkl = True
                            load_option = "default"
                        continue

                    # Confirm that file is .pkl
                    if preferences_file[-4:] != ".pkl":
                        messagebox.showerror(title="Invalid Extension", message="The selected file has the wrong "
                                                                                "extension. Please select a \".pkl\" "
                                                                                "file.")
                        preferences_file = ""
                        continue

                    # Confirm that unpickled file is a dictionary
                    with (open(preferences_file, "rb")) as pkl:
                        if not isinstance(pickle.load(pkl), dict):
                            messagebox.showerror(title="Invalid File Type",
                                                 message="The selected file cannot be unpickled to a dictionary. Please"
                                                         " try a different \".pkl\" file.")
                            preferences_file = ""
                            continue

                    # File is valid
                    is_pkl = True
                    load_option = "normal"

                # Ask user if file should be copied to expected location to avoid future prompts
                copy_file = messagebox.askyesno(title="Copy File", message="Should the Preferences file be copied to "
                                                                           "the expected location to avoid future "
                                                                           "prompts?")
                if copy_file:
                    shutil.copy2(preferences_file, f"{init_path}\\Preferences.pkl")

            # Create file
            else:
                use_default = messagebox.askyesno(title="Use Default?", message="Should the default preferences be "
                                                                                "used?")
                load_option = "default" if use_default else "user-select"

    # Load in from selected file
    if load_option == "normal" and preferences_file:
        with (open(preferences_file, "rb")) as pref:
            # Does not do direct assignment to ensure value of passed variable is updated
            preferences.update({key: value for key, value in pickle.load(pref).items()})

            # Update dictionary if any new items were added
            old_len = len(preferences.keys())
            for key, value in default_pref.items():
                if key not in preferences.keys():
                    preferences.update({key: value})

            # Show warning if keys were added
            if len(preferences.keys()) != old_len:
                messagebox.showwarning(title="Preferences Altered",
                                       message="The preferences dictionary was altered. This is either because a "
                                               "new key was added in a version change or because the preferences "
                                               "file did not contain the expected items.")

    # Load in defaults
    elif load_option == "default":
        preferences.update({key: value for key, value in default_pref.items()})

    # Load in by Edit Preferences
    elif load_option == "user-select":
        messagebox.showinfo(title="Select Preferences", message="Please set up your preferences in the following "
                                                                "window to continue.")

        # Use default preferences as starting point and set up appearance
        preferences.update({key: value for key, value in default_pref.items()})
        set_styles(preferences, win)

        # Launch Edit Preferences Frame
        epf = Toplevel(win)
        epf.title("Edit Preferences")
        new_pref = default_pref
        EditPreferencesFrame(epf, new_pref)
        epf.wait_window()


    return preferences_file


def add_shortcuts(preferences: dict, pref_file: str) -> None:
    """
    Check for existence of desktop and start menu shortcuts and create or edit shortcuts as needed.

    Imports win32com.client for shortcut creation. If win32com is not available, it can be automatically installed
    using pip. If the installation fails (or the system is not Windows), the shortcut prompt flag is the preferences
    dictionary will be set to False to avoid re-prompting the user each time the application is launched.

    :param preferences: Dictionary of program configuration options
    :param pref_file: Full path name of the preferences dictionary pickle file
    :return:
    """

    # Return if shortcuts should not be created and no shortcuts exist
    if (not preferences["Shortcut Prompt"] and preferences["Desktop Shortcut"] == "" and
            preferences["Start Menu Shortcut"] == ""):
        return

    exit_at_end = False  # Flag to specify whether application should be exited

    # Return if system is not Windows
    if platform.system().lower() != "windows":
        messagebox.showinfo(title="Auto-Shortcuts Unavailable", message="The automatic shortcut installation feature "
                                                                        "is only available on Windows. The core of the "
                                                                        "program can still be used.")
        preferences["Shortcut Prompt"] = False  # To avoid reprompt on next start


    # Import win32com.client and install if not available and user opts to auto-install it
    #   If import fails or if user opts to not install, can continue, but without ability to add shortcuts
    try:
        import win32com.client

    except ModuleNotFoundError:
        auto_inst = messagebox.askyesno(title="Missing win32com",
                                        message="The \"pywin32\" package was not found.\nWould you like to try to "
                                                "automatically install pywin32?\nNote: If this is the second time "
                                                "you've seen this message, auto-installation has failed. Please "
                                                "select No and install the package manually.")
        if auto_inst:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            messagebox.showinfo(title="Restart Required", message="A program restart is necessary to confirm the "
                                                                  "\"pywin32\" module was successfully installed.")
            preferences["Desktop Shortcut"] = "Check install"
            os.execv(sys.executable, ['python'] + [f"\"{sys.argv[0]}\""])


        else:
            cont = messagebox.askyesno(title="Continue?",
                                       message="You can still use the program without the win32com package, but "
                                               "shortcuts cannot be automatically created. Would you like to "
                                               "continue?")
            preferences["Shortcut Prompt"] = False  # To avoid reprompt on next start
            if not cont:
                exit_at_end = True

    # Save changes to preferences dictionary if shortcut cannot be created
    if not preferences["Shortcut Prompt"]:
        with (open(pref_file, "wb")) as pref:
            pickle.dump(preferences, pref)

        if not exit_at_end:
            return
        else:
            raise SystemExit

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

    # Get preferences and set up formatting
    preferences = {}
    pref_file = get_preferences(preferences, win)
    set_styles(preferences, win)

    # Set up shortcuts if needed
    add_shortcuts(preferences, pref_file)

    # Populate MainFrame
    MainFrame(win, preferences, pref_file)

    win.mainloop()


if __name__ == __main__():
    __main__()
