import copy
import os
import sys
from LabelButton import *
from tkinter import messagebox
from tkinter import ttk

initPath = os.path.dirname(sys.argv[0]) + "\\Files\\"
sys.path.append(initPath)


class EditPreferencesFrame:
    """A Tkinter window to edit program preferences and file locations. """
    def save_settings(self) -> None:
        """Store selected preferences setting in the preferences dictionary, then call on_close."""

        self.preference_dict["Font Type"] = self.font_select.get()

        # Ensure font size is a number and less than 20
        self.preference_dict["Font Size"] = self.size_select.get()
        try:
            if int(self.preference_dict["Font Size"]) > 20:
                self.preference_dict["Font Size"] = "20"
                messagebox.showinfo(title="Font Size Decreased",
                                    message="Maximum font size was exceeded. Reduced to size 20.")
        except ValueError:
            messagebox.showerror(title="Invalid Entry", message="Font size entered cannot be converted to an integer.\n"
                                                                "Please try again.")
            return

        if self.dark_mode_select.get():
            self.preference_dict["Dark Mode"] = True
        else:
            self.preference_dict["Dark Mode"] = False

        self.preference_dict["Combine Non-Sequential File Selections on Move"] = self.combine_non_seq_select.get()
        
        # Close window
        self.on_close()

    def revert_settings(self) -> None:
        """Reset preferences to original values."""

        # Prompt user to confirm they wish to revert changes
        confirm = messagebox.askyesno(title="Confirm Reversion",
                                      message="Are you sure you wish to revert settings to their previous values?")
        if not confirm:
            self.win.lift()
            return

        # Reset variable values
        self.font_select.set(self.orig_preference_dict["Font Type"])
        self.size_select.set(self.orig_preference_dict["Font Size"])
        self.dark_mode_select.set(self.orig_preference_dict["Dark Mode"])
        self.combine_non_seq_select.set(self.orig_preference_dict["Combine Non-Sequential File Selections on Move"])

        self.win.lift()

    def flip_shortcut(self) -> None:
        """Set the Shortcut Prompt flag to True and show message that shortcuts will be added on next launch."""
        self.preference_dict["Shortcut Prompt"] = True
        messagebox.showinfo(title="Shortcut Addition", message="You will be prompted to add shortcuts the next time "
                                                               "PDF Combiner is launched.")
        self.win.lift()

    def on_close(self) -> None:
        """Show info message box if needed, then close window."""

        # Unsaved change: Prompt if preferences should be saved before exiting
        if (self.preference_dict["Font Type"] != self.font_select.get() or
                int(self.preference_dict["Font Size"]) != int(self.size_select.get()) or
                self.preference_dict["Dark Mode"] != self.dark_mode_select.get() or
                self.preference_dict["Combine Non-Sequential File Selections on Move"] !=
                    self.combine_non_seq_select.get()):
            save = messagebox.askyesno(title="Unsaved Changes", message="You have unsaved changes. Would you like to "
                                                                        "save them before exiting?")
            if save:
                self.save_settings()
                return  # Returns to on_close at end of save_settings
            else:
                self.win.destroy()
                return

        # Changed appearance: Show note that will need to reload to be effective
        if (self.orig_preference_dict["Font Type"] != self.font_select.get() or
                int(self.orig_preference_dict["Font Size"]) != int(self.size_select.get()) or
                self.orig_preference_dict["Dark Mode"] != self.dark_mode_select.get()):
            messagebox.showinfo(title="Reload Necessary", message="Changes saved successfully.\n"
                                                                  "Application must be reloaded for appearance changes "
                                                                  "to take effect.")

        # Changed something other than appearance: Show message for successful completion
        elif (self.orig_preference_dict["Combine Non-Sequential File Selections on Move"] !=
                self.combine_non_seq_select.get()):
            messagebox.showinfo(title="Save Successful", message="Changes saved successfully.")

        # Close window
        self.win.destroy()

    def on_type(self, event: Event) -> None:
        """
        Input validation for font size entry.

        :param event: Event from the Key Press event handler
        """

        # If not in entry field, ignore
        if not isinstance(event.widget, ttk.Entry):
            return

        # Definition of allowable characters
        allowed_chars = [8, 9, 13, 16, 17, 18, 37, 38, 39, 40, 46]
        for i in range(48, 58):
            allowed_chars.append(i)
        for i in range(96, 106):
            allowed_chars.append(i)

        if (event.keycode not in allowed_chars) or len(self.size_select.get()) > 2:
            position = self.size_box.index(INSERT)
            self.size_select.set(self.size_select.get()[:position - 1] + self.size_select.get()[position:])
            self.size_box.icursor(position - 1)

    def __init__(self, win: Toplevel, preferences: dict) -> None:
        """
        Fill in the Toplevel window to edit program preferences and file locations.

        :param win: Toplevel window in which to create frame
        :param preferences: Dictionary of user program preferences
        """

        self.win = win
        self.preference_dict = preferences
        self.orig_preference_dict = copy.deepcopy(preferences)
        self.dark_mode = self.preference_dict["Dark Mode"]
        self.font_type = self.preference_dict["Font Type"]
        self.font_size = self.preference_dict["Font Size"]
        self.combine_non_seq = self.preference_dict["Combine Non-Sequential File Selections on Move"]

        if self.dark_mode:
            self.win.config(bg="black")

        # Font details
        self.font_frame = ttk.Frame(self.win)
        self.font_frame.grid(row=0, column=0, padx=5, pady=1, sticky="ew")

        # Font style
        ttk.Label(self.font_frame, text="Font:").grid(row=0, column=0, padx=(5, 1), pady=5, sticky="w")
        self.font_select = StringVar()
        font_options = ["Arial", "Calibri", "Cambria", "Times New Roman"]
        self.font_select.set(self.preference_dict["Font Type"])

        if self.font_select.get() not in font_options:
            messagebox.showwarning(title="Font Not Available", message="An invalid font was found in the file.\n"
                                                                       "Font was changed to default.")
            self.font_select.set("Times New Roman")
        self.font_list = ttk.Combobox(self.font_frame, values=font_options, textvariable=self.font_select,
                                      state="readonly")
        self.font_list.grid(row=0, column=1, padx=(1, 2), pady=5, sticky="w")
        
        # Font size
        ttk.Label(self.font_frame, text="Font Size:").grid(row=0, column=2, padx=(2, 1), pady=5, sticky="w")
        self.size_select = StringVar(self.win)
        self.size_select.set(self.preference_dict["Font Size"])
        self.size_box = ttk.Spinbox(self.font_frame, from_=8, to=20, textvariable=self.size_select)
        self.size_box.grid(row=0, column=3, padx=(1, 5), pady=5, sticky="w")
        self.size_box.set(self.preference_dict["Font Size"])

        # Dark mode
        self.dark_mode_frame = ttk.Frame(self.win)
        self.dark_mode_frame.grid(row=1, column=0, padx=5, pady=1, sticky="ew")

        ttk.Label(self.dark_mode_frame, text="Dark Mode:").grid(row=0, column=0, padx=(5, 1), pady=5, sticky="e")
        self.dark_mode_select = BooleanVar(value=self.preference_dict["Dark Mode"])
        self.dark_mode_box = ttk.Checkbutton(self.dark_mode_frame, variable=self.dark_mode_select)
        self.dark_mode_box.grid(row=0, column=1, padx=(1, 5), pady=5, sticky="w")

        # Combine non-sequential files
        self.non_seq_frame = ttk.Frame(self.win)
        self.non_seq_frame.grid(row=2, column=0, padx=5, pady=1, sticky="ew")

        (ttk.Label(self.non_seq_frame, text="Combine Non-Sequential File Selections on Move:")
         .grid(row=0, column=0, columnspan=3, padx=(5, 1), pady=5, sticky="w"))
        self.combine_non_seq_select = StringVar(value=self.combine_non_seq)
        self.combine_non_seq_radio = ttk.Radiobutton(self.non_seq_frame, text="Always",
                                                     variable=self.combine_non_seq_select, value="Always")
        self.combine_non_seq_radio.grid(row=1, column=0, padx=1, pady=5, sticky="w")

        self.combine_non_seq_radio = ttk.Radiobutton(self.non_seq_frame, text="Ask",
                                                     variable=self.combine_non_seq_select, value="Ask")
        self.combine_non_seq_radio.grid(row=1, column=1, padx=1, pady=5, sticky="w")

        self.combine_non_seq_radio = ttk.Radiobutton(self.non_seq_frame, text="Never",
                                                     variable=self.combine_non_seq_select, value="Never")
        self.combine_non_seq_radio.grid(row=1, column=2, padx=1, pady=5, sticky="w")

        # Action buttons
        self.btn_frame = ttk.Frame(self.win)
        self.btn_frame.grid(row=3, column=0, padx=5, pady=1, sticky="ew")

        (LabelButton(self.btn_frame, text="Save", command=lambda: self.save_settings(), dark_mode=self.dark_mode)
         .grid(row=0, column=0, padx=5, pady=5, sticky="e"))
        (LabelButton(self.btn_frame, text="Revert", command=lambda: self.revert_settings(), dark_mode=self.dark_mode)
         .grid(row=0, column=1, padx=5, pady=5, sticky="w"))
        (LabelButton(self.btn_frame, text="Add Shortcuts", command=lambda: self.flip_shortcut(),
                     dark_mode=self.dark_mode)
         .grid(row=0, column=2, padx=5, pady=5, sticky="w"))

        self.win.protocol("WM_DELETE_WINDOW", lambda: self.on_close())

        self.win.bind("<KeyPress>", lambda e: self.on_type(e))

        # Move on top of main window
        self.win.update_idletasks()
        self.win.geometry(f"{self.win.winfo_reqwidth()}x{self.win.winfo_reqheight()}+{self.win.master.winfo_x()}+"
                          f"{self.win.master.winfo_y()}")
