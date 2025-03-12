from ScrollableFrame import *
from EditPreferencesFrame import *
from PageSelection import *
from Tooltip import Tooltip
from tkinter import messagebox, filedialog
import os
import sys
import copy
import subprocess
from threading import Thread
from datetime import datetime
import pickle
import time

init_path = os.path.dirname(sys.argv[0]) + "\\Files"
sys.path.append(init_path)

# Import pypdf and install if not available and user opts to auto-install it
try:
    from pypdf import PdfWriter, PdfReader
except ModuleNotFoundError:
    auto_install = messagebox.askyesno(title="Missing pypdf", message="The \"pypdf\" package was not found.\nWould you "
                                                                      "like to try to automatically install pypdf?")
    if auto_install:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
        from pypdf import PdfWriter, PdfReader
        messagebox.showinfo(title="pypdf Installed", message="The pypdf package was successfully installed.")
    else:
        messagebox.showerror(title="Missing pypdf", message="Please install the pypdf package.")
        raise ModuleNotFoundError

def get_path_name(extension: str) -> str:
    """
    Prompt the user to select a file with the given extension.

    :param extension: File extension for which to search, including the dot
    :return: Full pathname as a string or an empty string if no file was selected
    """

    path = ""
    while path == "":
        if extension == ".txt":
            path = filedialog.askopenfilename(initialdir=init_path, filetypes=[("Text Files", "*.txt")])
        elif extension == ".pdf":
            path = filedialog.askopenfilename(initialdir=init_path, filetypes=[("PDF Files", "*.pdf")])
        else:
            path = filedialog.askopenfilename(initialdir=init_path)

        # File was not chosen: ask if user wishes to retry
        if path == "":
            confirm = messagebox.askyesno(title="No File Selected",
                                          message="No file was selected. Would you like "
                                                  "to retry?")
            if not confirm:
                return ""
            continue

        # Confirm that file extension matches parameter
        if path[-len(extension):] != extension:
            messagebox.showerror(title="Invalid Extension",
                                 message=f"A file with an invalid extension was chosen. "
                                         f"Please select a \"{extension}\" file.")
            path = ""

    return path

class MainFrame:

    # File selection frame creation
    def __init__(self, win: Tk, preferences: dict, pref_file: str) -> None:
        """
        Populate main frame with file selection and order manipulation elements.

        :param win: Tk object in which to draw frame
        :param preferences: Dictionary of program preferences
        :param pref_file: Path name of preferences file
        """

        # Passed parameters
        self.win = win
        self.preferences = preferences
        self.dark_mode = preferences["Dark Mode"]
        self.font_type = preferences["Font Type"]
        self.font_size = int(preferences["Font Size"])
        self.pref_file = pref_file

        # Other parameters
        self.header_str = "Start adding files using the frame above."
        self.file_info = []  # List of (full path, file name) tuples
        self.checkboxes = []
        self.check_states = []
        self.placeholder_deleted = False
        self.selected_indices = []
        self.selected_pages = {}  # Stores info from the page_sel window in the right-click event handler
        self.write_pages = {}  # Stores "cleaned" info from the selected_pages dictionary

        # Merger frame attributes
        self.is_writing = False
        self.files_writen = False
        self.merger_frame = None
        self.merger = None
        self.save_path = ""
        self.total_size = 0

        # Populate main frame
        self.win.iconify()

        #   File selection frame
        self.selections_frame = ttk.Labelframe(self.win, text="File Selector")
        self.selections_frame.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        #       File path name manual entry
        ttk.Label(self.selections_frame, text="File Path:").grid(row=0, column=0, padx=(5, 1), pady=5, sticky="e")
        self.path_entered = StringVar()
        self.path_entry = ttk.Entry(self.selections_frame, width=60, textvariable=self.path_entered,
                               font=(self.font_type, self.font_size - 1))
        self.path_entry.grid(row=0, column=1, columnspan=2, padx=(1, 5), pady=(5, 1), sticky="w")

        self.invalid_char_label = ttk.Label(self.selections_frame, text="")  # Label for when invalid char is typed
        self.invalid_char_label.grid(row=1, column=1, columnspan=2, padx=5, pady=(1, 5), sticky="w")

        #       Action buttons
        self.add_multi = LabelButton(self.selections_frame, text="Add Multiple", dark_mode=self.dark_mode,
                                     command=lambda: self.add_files("multi"))
        self.add_multi.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        Tooltip(self.add_multi, text="Add multiple PDF files using the file dialog box")

        self.load_existing = LabelButton(self.selections_frame, text="Load Saved List", dark_mode=self.dark_mode,
                                         command=lambda: self.load_files())
        self.load_existing.grid(row=2, column=1, padx=5, pady=5, sticky="e")
        self.add_single = LabelButton(self.selections_frame, text="Add", dark_mode=self.dark_mode,
                                      command=lambda: self.add_files("single"), state="disabled")
        self.add_single.grid(row=2, column=2, padx=5, pady=5, sticky="e")

        #   File order frame
        self.selected_frame = ttk.Labelframe(self.win, text="Files Selected")
        self.selected_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        #       Header items
        border_frame = ttk.Frame(self.selected_frame)
        border_frame.grid(row=0, column=0, padx=5, pady=1, sticky="w")
        self.index_header = ttk.Label(border_frame, text="      0:", style="Hidden.TLabel")
        self.index_header.grid(row=0, column=0, padx=(1.5, 1), pady=0.5, sticky="w")

        self.all_boxes = ttk.Checkbutton(border_frame, command=lambda: self.all_boxes_selection())
        self.all_boxes.grid(row=0, column=1, padx=1, pady=1, sticky="w")
        self.all_boxes.state(["!selected", "!alternate"])

        Tooltip(self.all_boxes, text="Select or deselect all listed files")

        ttk.Label(border_frame, text="File Names:").grid(row=0, column=2, padx=(1, 5), pady=1, sticky="w")

        self.scroll_frame = ScrollableFrame(self.selected_frame, dark_mode=self.dark_mode, location=[0, 1], span=[1, 1],
                                            dims="y")

        #       Placeholders
        style = ttk.Style()
        style.configure("Hidden.TLabel", foreground=str(self.win.cget("bg")), background=str(self.win.cget("bg")))
        ttk.Label(self.scroll_frame, text="      0:", style="Hidden.TLabel").grid(row=0, column=0, padx=(5, 1),
                                                                                  pady=0.5, sticky="e")
        blank = ttk.Checkbutton(self.scroll_frame)
        blank.grid(row=0, column=1, padx=1, pady=0.5)
        blank.state(["!selected", "!alternate"])

        ttk.Label(self.scroll_frame, text=self.header_str).grid(row=0, column=2, padx=(1, 5), pady=0.5)

        #       File count
        self.file_count = ttk.Label(self.selected_frame, text="  0 files selected (0.0 MB)")
        self.file_count.grid(row=7, column=0, padx=5, pady=1, sticky="w")

        #       Adding blank pages between files
        self.add_blank_page = BooleanVar(value=self.preferences["Add Blank Page Between Files"])
        self.add_blank_page_box = ttk.Checkbutton(self.selected_frame, text="Add a blank page between files",
                                                  variable=self.add_blank_page, command=self.toggle_blank_page)
        self.add_blank_page_box.grid(row=8, column=0, padx=5, pady=5, sticky="w")

        #       Movement buttons
        btn_frame = ttk.Frame(self.selected_frame)
        btn_frame.grid(row=1, column=1, padx=5, pady=5, sticky="ne")

        style.configure("Arrow.TLabel", foreground=str(self.win.cget("bg")), background=str(self.win.cget("bg")),
                        font=(preferences["Font Type"], 20))

        self.move_top = LabelButton(btn_frame, text="\u219F", style="Arrow.TLabel", dark_mode=self.dark_mode,
                                    state="disabled", command=lambda: self.move_files("top"))
        self.move_top.grid(row=2, column=1, padx=5, pady=5, ipadx=8.5, sticky="n")
        self.move_up = LabelButton(btn_frame, text="\u2B61", style="Arrow.TLabel", dark_mode=self.dark_mode,
                                   state="disabled", command=lambda: self.move_files("up"))
        self.move_up.grid(row=3, column=1, padx=5, pady=5, ipadx=8.5, sticky="n")
        self.move_down = LabelButton(btn_frame, text="\u2B63", style="Arrow.TLabel", dark_mode=self.dark_mode,
                                     state="disabled", command=lambda: self.move_files("down"))
        self.move_down.grid(row=4, column=1, padx=5, pady=5, ipadx=8.5, sticky="n")
        self.move_bottom = LabelButton(btn_frame, text="\u21A1", style="Arrow.TLabel", dark_mode=self.dark_mode,
                                       state="disabled", command=lambda: self.move_files("bottom"))
        self.move_bottom.grid(row=5, column=1, padx=5, pady=5, ipadx=8.5, sticky="n")
        # Unicode codes: 2B61: up arrow; 2B63 down arrow; 219F double up arrow; 21A1 double down arrow; 21A5/7 w/ bar

        #       Delete button
        # Unicode 1F5D1 = 128465 decimal
        self.delete = LabelButton(btn_frame, text=chr(128465), style="Arrow.TLabel", dark_mode=self.dark_mode,
                                  state="disabled", command=lambda: self.remove_files("selected"))
        self.delete.grid(row=6, column=1, padx=5, pady=5, sticky="n")

        #       Button tooltips
        Tooltip(self.move_top, text="Move selected files to top of list")
        Tooltip(self.move_up, text="Move selected files up one position")
        Tooltip(self.move_down, text="Move selected files down one position")
        Tooltip(self.move_bottom, text="Move selected files to bottom of list")
        Tooltip(self.delete, text="Remove selected files from list")

        #   Overall Buttons
        button_frame = ttk.Frame(self.win)
        button_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.cancel = LabelButton(button_frame, text="Clear", dark_mode=self.dark_mode,
                                  command=lambda: self.remove_files("all"))
        self.cancel.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        Tooltip(self.cancel, text="Remove all files from list")

        self.save = LabelButton(button_frame, text="Save List", dark_mode=self.dark_mode, state="disabled",
                                command=lambda: self.save_files())
        self.save.grid(row=0, column=1, padx=5, pady=5)
        self.next = LabelButton(button_frame, text="Merge", dark_mode=self.dark_mode, state="disabled",
                                command=lambda: self.merge_files())
        self.next.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.edit_pref = LabelButton(button_frame, text="Preferences", dark_mode=self.dark_mode,
                                     command=lambda: self.edit_preferences())
        self.edit_pref.grid(row=0, column=3, padx=5, pady=5, sticky="e")

        # Virtual binds
        self.win.bind("<KeyPress>", lambda e: self.on_type(e.widget))
        self.win.unbind("<Return>")
        self.win.bind("<Return>", lambda e: self.add_files("single"))
        self.win.bind("<Button-1>", lambda e: self.on_click(e.widget))
        self.win.bind("<Button-3>", lambda e: self.on_right_click(e.widget))

        # Finish window configuration
        self.win.rowconfigure(index=1, weight=1)
        self.win.update_idletasks()
        self.win.deiconify()
        self.win.minsize(self.win.winfo_reqwidth(), self.win.winfo_reqheight())
        self.min_size = (self.win.winfo_reqwidth(), self.win.winfo_reqheight())
        self.win.maxsize(self.win.winfo_reqwidth(), self.win.winfo_screenheight() - 80)
        self.win.protocol("WM_DELETE_WINDOW", self.on_close)

        self.path_entry.focus_set()

    # File selection frame methods

    #   Virtual bind functions
    def on_type(self, widget: tkinter.Widget) -> None:
        """Remove invalid character entries and enable or disable the "Add" button if in path entry widget."""

        if widget != self.path_entry:
            return

        # Clear invalid character message
        def del_message() -> None:
            self.invalid_char_label.configure(text="")

        # Remove invalid characters (\ : * ? " < > |) and show message
        #   Special case for colon: can be in second position for drive letter
        try:
            input_str = self.path_entered.get()
            new_str = ""
            invalid_char = ""
            for i, c in enumerate(input_str):
                if c not in ["\\", "*", "?", "\"", "<", ">", "|", ":"] or (c == ":" and i == 1):
                    new_str += c
                else:
                    invalid_char = c

            if len(new_str) < len(input_str):
                if self.path_entry.index(INSERT) != len(input_str):
                    self.path_entry.icursor(self.path_entry.index(INSERT) - 1)
                else:
                    self.path_entry.icursor(self.path_entry.index(INSERT))
                self.path_entered.set(new_str)
                self.invalid_char_label.configure(text=f"\"{invalid_char}\" is not a valid path character")
                self.invalid_char_label.after(ms=2000, func=del_message)

        except IndexError:  # Empty string
            pass

        if len(self.path_entered.get()) > 0:
            self.add_single.set_state("normal")
        else:
            self.add_single.set_state("disabled")

    def on_click(self, widget: tkinter.Widget) -> None:
        """
        Toggle the checkbox for the selected file and update the selected file list

        :param widget: Widget which triggered the mouse event
        :return:
        """

        # Pass if widget is not in the scrollable frame
        if "scrollableframe" not in str(widget):
            return

        # Pass if no files have been added
        if len(self.file_info) == 0:
            return

        # Determine widget index number
        try:
            index = widget.grid_info()["row"]
        except KeyError:  # Clicked in area near widget that was not a widget
            return

        # Remove index from list if box was unticked
        if self.check_states[index].get():
            for position, i in enumerate(self.selected_indices):
                if index == i:
                    self.selected_indices.pop(position)
                    break

        # Add index to list if box was ticked
        else:
            self.selected_indices.append(index)
            self.selected_indices.sort()

        # Toggle checkbox state if clicked on label
        if isinstance(widget, ttk.Label):
            self.check_states[index].set(not self.check_states[index].get())

        self.update_widgets()

    def on_right_click(self, widget: tkinter.Widget) -> None:
        """
        Determine index of selected widget, create selected_pages entry if necessary, then create a PageSelection
        window.

        :param widget: Widget which triggered the mouse event
        :return:
        """

        # Pass if widget is not in the scrollable frame
        if "scrollableframe" not in str(widget):
            return

        # Pass if file_info is empty
        if len(self.file_info) == 0:
            return

        # Determine widget index number
        try:
            index = widget.grid_info()["row"]
        except KeyError:  # Clicked in area near widget that was not a widget
            return

        # Check if file is already in page selection dictionary
        if self.file_info[index][0] not in self.selected_pages.keys():
            self.selected_pages.update({self.file_info[index][0]: ("", True, True)})

        # Call page selection window
        page_sel = PageSelection(self.win, self.preferences, index, self.file_info[index], self.selected_pages)

    #   File handling
    def add_files(self, mode: str = ("multi", "single"), prompt: bool = True) -> None:
        """
        Add selected file(s) to list and update window.

        If in multiple selection mode, calls filedialog box first is prompt is true. Deletes placeholders if necessary.

        :param mode: File addition type
        :param prompt: Flag for whether the list of files should be prompted from user
        :return:
        """

        # Temporary list to store file additions
        add_list = []

        # Add multiple (with prompt): open file dialog
        if mode == "multi" and prompt:
            path_list = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])

            # Add files to appropriate list
            skipped_files = []
            for path in path_list:
                if path[-4:] != ".pdf":
                    skipped_files.append(path)
                else:
                    add_list.append((path, path.split("/")[-1]))

            # Show messagebox if files were not added
            if len(skipped_files) > 0:
                message = "The following files were not PDF files and will not be added to the list:\n"
                for skipped_file in skipped_files:
                    message += f"{skipped_file.split('/')[-1]}\n"
                messagebox.showwarning(title="Skipped Files", message=message)

        # Add multiple (without prompt): Use existing file_info
        elif mode == "multi" and not prompt:
            add_list = copy.deepcopy(self.file_info)
            self.file_info.clear()

        # Add single: Check if inputted file exists
        elif mode == "single":
            # Return if main window is not in focus
            if not self.win.focus_get():
                return
            path_name = self.path_entered.get()
            path_name = path_name.replace("\\", "/")
            # Append PDF extension if needed
            if path_name[-4:] != ".pdf":
                path_name += ".pdf"

            if not os.path.exists(fr"{path_name}"):
                messagebox.showerror(title="File Not Found", message=f"The file \"{path_name}\" could not be found. "
                                                                     f"Please try a different path.")
                return
            else:
                add_list.append((path_name, path_name.split("/")[-1]))

            self.path_entered.set("")

        # Empty file list: return
        if len(add_list) == 0:
            return

        # Delete placeholders if needed
        if not self.placeholder_deleted:
            for widget in self.scroll_frame.winfo_children():
                widget.destroy()
            self.placeholder_deleted = True

            # Add original file name placeholder text as "hidden" header to preserve width
            self.scroll_frame.add_header(column=2, text=self.header_str)

        # Update window accordingly
        for info in add_list:
            # Add checkbutton to list
            self.check_states.append(BooleanVar(value=False))
            self.checkboxes.append(ttk.Checkbutton(self.scroll_frame, variable=self.check_states[-1]))

            # Add file details to main list
            self.file_info.append(info)

            # Add details to frame
            row = len(self.file_info) - 1
            ttk.Label(self.scroll_frame, text=f"{row + 1:> 7}:").grid(row=row, column=0, padx=(5, 1), pady=0.5,
                                                                      sticky="e")
            self.checkboxes[-1].grid(row=row, column=1, padx=1, pady=0.5)
            test = ttk.Label(self.scroll_frame, text=info[1], cursor="hand2")
            test.grid(row=row, column=2, padx=(1, 5), pady=0.5, sticky="w")
            Tooltip(test, text=f"Full path: {info[0]}\nRight click to select pages")

        self.update_widgets()

    def load_files(self) -> None:
        """Load the list of files from the user-specified list. Obtain new file location if needed."""

        load_count = 0  # Number of files successfully located

        # Prompt user for save file location
        messagebox.showinfo(title="Select Save", message="Select the text file from which to load the list of files "
                                                         "on the next screen.")
        load_file = get_path_name(".txt")
        if load_file == "":  # User did not select file
            return

        # Populate file_info dictionary; ask for revised file locations as necessary
        with (open(load_file, "r")) as load:
            lines = load.readlines()

        for line in lines:
            line = line.strip()
            dict_include = False  # Flag for whether a dictionary entry should be added to page_selections

            # Separate line into components
            try:
                path_name, pages_sel, resort, remove_dup = line.split("\t")
                dict_include = True
            except ValueError:  # File did not have page selection details associated with it
                path_name = line
                pages_sel = ""
                resort = ""
                remove_dup = ""

            # Clean up path name
            path_name = path_name.replace("\\", "/")
            file_name = path_name.split("/")[-1]

            if file_name == "":  # Skip blank lines
                continue

            # Ask for revised path if file does not exist
            if not os.path.exists(path_name):
                revise = messagebox.askyesno(title="File Not Found", message=f"The file \"{file_name}\" was not found. "
                                                                             f"Would you like to replace the file with "
                                                                             f"another?")

                if revise:
                    path_name = get_path_name(".pdf")
                    file_name = path_name.split("/")[-1]
                else:
                    path_name = ""

            # Add info to dictionary (if file exists)
            if path_name != "":
                self.file_info.append((path_name, file_name))
                load_count += 1

            # Add details to dictionary (if file exists)
            if dict_include and path_name != "":
                resort = True if resort == "True" else False
                remove_dup = True if remove_dup == "True" else False
                self.selected_pages.update({path_name: (pages_sel, resort, remove_dup)})

        # Add files to window
        self.add_files(mode="multi", prompt=False)

        # Show completion message
        file_component = f"1 file was" if load_count == 1 else f"{load_count} files were"
        messagebox.showinfo(title="Load Successful", message=f"{file_component} successfully loaded.")

    def move_files(self, move: str = ("top", "up", "down", "bottom")) -> None:
        """
        Move selected files as specified.

        Treatment of non-sequential index numbers is set by user preference.

        :param move: String specifying which file movement should take place
        :return:
        """

        # Determine if non-sequential index numbers are present
        non_seq = False

        for i, index in enumerate(self.selected_indices):
            try:
                if self.selected_indices[i + 1] - index > 1:
                    non_seq = True
                    break
            except IndexError:
                pass

        # Determine whether non-sequential files should be combined
        combine_non_seq = False
        if non_seq and self.preferences["Combine Non-Sequential File Selections on Move"] == "Always":
            combine_non_seq = True
        elif non_seq and self.preferences["Combine Non-Sequential File Selections on Move"] == "Ask":
            combine_non_seq = messagebox.askyesno(title="Combine Non-Sequential Files",
                                                  message="Files with non-sequential indices were selected. Should "
                                                          "these files be moved next to each other after moving?")

        # Move to top
        if move == "top":
            # Files are sequentially selected (or should be combined)
            if not non_seq or combine_non_seq:
                for i, index in enumerate(self.selected_indices):
                    self.file_info.insert(i, self.file_info[index])
                    self.file_info.pop(index + 1)

            # Files are not sequential and should not be combined
            elif non_seq and not combine_non_seq:
                move = self.selected_indices[0]  # All files should be moved up by this amount
                for index in self.selected_indices:
                    move_info = self.file_info[index]
                    self.file_info.pop(index)
                    self.file_info.insert(index - move, move_info)

        # Move up
        elif move == "up":
            # Files are sequentially selected (or should not be combined)
            if not non_seq or (non_seq and not combine_non_seq):
                for i, index in enumerate(self.selected_indices):
                    if index == 0:
                        continue
                    self.file_info.insert(index - 1, self.file_info[index])
                    self.file_info.pop(index + 1)

            # Files are not sequential and should be combined
            elif non_seq and combine_non_seq:
                for i, index in enumerate(self.selected_indices):
                    if index == 0:
                        continue
                    info = self.file_info[index]
                    self.file_info.pop(index)
                    self.file_info.insert(i + self.selected_indices[0], info)

        # Move down
        elif move == "down":
            max_index = len(self.file_info)
            # Files are sequentially selected (or should not be combined)
            if not non_seq or (non_seq and not combine_non_seq):
                for i, index in enumerate(reversed(self.selected_indices)):
                    if index == max_index:
                        continue
                    info = self.file_info[index]
                    self.file_info.pop(index)
                    self.file_info.insert(index + 1, info)

            # Files are not sequential and should be combined
            elif non_seq and combine_non_seq:
                for i, index in enumerate(reversed(self.selected_indices)):
                    if index == max_index:
                        continue
                    info = self.file_info[index]
                    self.file_info.insert(self.selected_indices[-1] + 1 - i, info)
                    self.file_info.pop(index)

        # Move to bottom
        elif move == "bottom":
            # Files are sequentially selected (or should be combined)
            if not non_seq or combine_non_seq:
                for i, index in enumerate(reversed(self.selected_indices)):
                    self.file_info.insert(len(self.file_info) - i, self.file_info[index])
                    self.file_info.pop(index)


            # Files are not sequential and should not be combined
            elif non_seq and not combine_non_seq:
                move = len(self.file_info) - self.selected_indices[-1] - 1  # All files should be moved down by this amt
                for index in reversed(self.selected_indices):
                    move_info = self.file_info[index]
                    self.file_info.pop(index)
                    self.file_info.insert(index + move, move_info)

        # Redraw window
        #   Delete existing file labels
        for widget in self.scroll_frame.winfo_children():
            # Only delete if in column 2 (file name label) and is not header label
            if widget.grid_info()["column"] == 2 and self.header_str not in widget.cget("text"):
                widget.destroy()

        #   Add new file labels
        for i, (_, file_name) in enumerate(self.file_info):
            ttk.Label(self.scroll_frame, text=file_name).grid(row=i, column=2, padx=(1, 5), pady=0.5, sticky="w")

        #   Clear all checkboxes
        for state in self.check_states:
            state.set(False)

        # Reset selected indices
        self.selected_indices = []

        self.update_widgets()

    def remove_files(self, mode: str = ("all", "selected")) -> None:
        """
        Remove specified files after prompt.

        :param mode: String for mode type; "all" to delete all files, "selected" to only delete selected files
        :return:
        """

        # Delete all: clear file and checkbutton lists, put back in placeholders
        if mode == "all":
            confirm = messagebox.askyesno(title="Confirm Deletion", message="Are you sure you want to delete all files "
                                                                           "from the list?")
            if not confirm:
                return

            # Clear lists
            self.file_info.clear()
            self.checkboxes.clear()
            self.check_states.clear()
            self.selected_indices.clear()

            # Delete existing file labels
            for widget in self.scroll_frame.winfo_children():
                # Only delete if is not header label
                if self.header_str not in widget.cget("text"):
                    widget.destroy()

            # Delete page number entries
            self.selected_pages.clear()

        # Delete selected: clear selected files, reduce length of checkbutton lists
        elif mode == "selected":
            confirm = messagebox.askyesno(title="Confirm Deletion", message="Are you sure you want to delete the "
                                                                            "selected files from the list?")
            if not confirm:
                return

            # Remove selected files
            for i, index in enumerate(self.selected_indices):
                self.selected_pages[self.file_info[index - i][0]] = ("", True, True)
                self.file_info.pop(index - i)
                self.checkboxes[-1].grid_forget()
                self.checkboxes.pop()
                self.check_states.pop()

            # Redraw window
            #   Delete existing file labels
            for widget in self.scroll_frame.winfo_children():
                try:
                    # Only delete if in column 2 (file name label) and is not header label
                    if widget.grid_info()["column"] == 2 and self.header_str not in widget.cget("text"):
                        widget.destroy()

                    # Delete unneeded index labels
                    elif widget.grid_info()["column"] == 0 and widget.grid_info()["row"] >= len(self.file_info):
                        widget.destroy()

                except KeyError:
                    pass

            #   Add new file labels
            for i, (file_path, file_name) in enumerate(self.file_info):
                label = ttk.Label(self.scroll_frame, text=file_name, cursor="hand2")
                label.grid(row=i, column=2, padx=(1, 5), pady=0.5, sticky="w")
                Tooltip(label, text=f"Full path: {file_path}\nRight click to select pages")

        #   Clear all checkboxes
        for state in self.check_states:
            state.set(False)

        self.selected_indices.clear()

        # Redraw placeholders if no files are in list
        if len(self.file_info) == 0:
            ttk.Label(self.scroll_frame, text="      0:", style="Hidden.TLabel").grid(row=0, column=0, padx=(5, 1),
                                                                                      pady=0.5, sticky="e")
            blank = ttk.Checkbutton(self.scroll_frame)
            blank.grid(row=0, column=1, padx=1, pady=0.5)
            blank.state(["!selected", "!alternate"])
            ttk.Label(self.scroll_frame, text=self.header_str).grid(row=0, column=2, padx=(1, 5), pady=0.5)

            self.placeholder_deleted = False

        self.update_widgets()

    def save_files(self, prompt: bool = True) -> None:
        """
        Prompt for a save location, then save the list of full path names as a text file.

        :param prompt: Flag for whether the user should be asked if the program should be terminated after saving
        :return:
        """

        # Prompt for save file location
        messagebox.showinfo(title="Select Save Location", message="Select the location in which to save the file on "
                                                                  "the next screen.")
        save_file = ""
        while save_file == "":
            save_file = filedialog.asksaveasfilename(initialdir=init_path)
            if save_file == "":
                confirm = messagebox.askyesno(title="No File Selected", message="No file was selected. Would you like "
                                                                                "to retry?")
                if not confirm:
                    if prompt:
                        return
                    self.win.destroy()
                    return

        # Save full path names to file
        if save_file[-4:] != ".txt":
            save_file += ".txt"

        with open(f"{save_file}", "w+") as file:
            for path, _ in self.file_info:
                file.write(f"{path}")
                try:
                    for item in self.selected_pages[path]:
                        file.write(f"\t{item}")
                except KeyError:  # Selected pages window was not launched for this file
                    pass
                file.write("\n")

        # Ask if program should be terminated
        if prompt:
            file_count = "1 file name was" if len(self.file_info) == 1 else f"{len(self.file_info)} file names were"
            close = messagebox.askyesno(title="Close Program?", message=f"{file_count} successfully saved to "
                                                                        f"\"{save_file}.\"\n"
                                                                        f"Would you like to exit the program?")
            if close:
                self.win.destroy()

        else:
            file_count = "1 file name was" if len(self.file_info) == 1 else f"{len(self.file_info)} file names were"
            messagebox.showinfo(title="Save Successful",
                                message=f"{file_count} successfully saved to \"{save_file}.\".")
            self.win.destroy()

    #   Window "handling"
    def update_widgets(self) -> None:
        """Update states of all buttons except add buttons and move window if needed."""

        # Change state of movement buttons
        if len(self.selected_indices) == 0:
            self.move_top.set_state("disabled")
            self.move_up.set_state("disabled")
            self.move_down.set_state("disabled")
            self.move_bottom.set_state("disabled")
            self.delete.set_state("disabled")
        else:
            self.move_top.set_state("normal")
            self.move_up.set_state("normal")
            self.move_down.set_state("normal")
            self.move_bottom.set_state("normal")
            self.delete.set_state("normal")

        # Change state of finish button
        if len(self.file_info) == 0:
            self.next.set_state("disabled")
            self.save.set_state("disabled")
            self.load_existing.grid()
        else:
            self.next.set_state("normal")
            self.save.set_state("normal")
            self.load_existing.grid_remove()

        # Change state of all_boxes
        total_selected = len(self.selected_indices)

        if total_selected == 0:
            self.all_boxes.state(["!selected", "!alternate"])
        elif total_selected == len(self.file_info):
            self.all_boxes.state(["selected", "!alternate"])
        else:
            self.all_boxes.state(["!selected", "alternate"])

        # Update file count label
        total_size = sum([os.path.getsize(file) for (file, _) in self.file_info]) / 1024 ** 2
        if len(self.file_info) == 1:
            self.file_count.configure(text=f"  1 file selected ({total_size:.1f} MB)")
        else:
            self.file_count.configure(text=f"{len(self.file_info): >3} files selected ({total_size:.1f} MB)")

        # Update placeholder label
        if len(self.file_info) < 10:
            self.index_header.configure(text="      0:")
        elif len(self.file_info) < 100:
            self.index_header.configure(text="     00:")
        elif len(self.file_info) < 1000:
            self.index_header.configure(text="    000:")
        else:
            self.index_header.configure(text="   0000:")

        # Resize and move window (if needed)
        self.win.update_idletasks()
        w = min(self.win.winfo_screenwidth(), self.win.winfo_reqwidth())
        h = min(self.win.winfo_screenheight() - 80, self.win.winfo_reqheight())
        x = max(0, min(self.win.winfo_x(), self.win.winfo_screenwidth() - self.win.winfo_reqwidth()))
        y = max(0, min(self.win.winfo_y(), self.win.winfo_screenheight() - self.win.winfo_reqheight() - 80))
        self.win.maxsize(self.win.winfo_reqwidth(), self.win.winfo_screenheight() - 80)
        self.win.geometry(f"{w}x{h}+{x}+{y}")

    def toggle_blank_page(self) -> None:
        """Store the current value of the add_blank_page checkbox to the preferences dictionary."""
        self.preferences["Add Blank Page Between Files"] = self.add_blank_page.get()

    def all_boxes_selection(self) -> None:
        """Change state of all checkboxes (and selected_indices list) based on state of all_boxes checkbox."""

        # Box is selected: Select all boxes
        if "selected" in self.all_boxes.state():
            for i, state in enumerate(self.check_states):
                state.set(True)
                if i not in self.selected_indices:
                    self.selected_indices.append(i)
            self.selected_indices.sort()

        # Box is not selected: Clear all boxes
        else:
            for i, state in enumerate(self.check_states):
                state.set(False)
            self.selected_indices.clear()

        self.update_widgets()

    #   Calling other frames
    def edit_preferences(self) -> None:
        pref_win = Toplevel(self.win)
        pref_win.title("Preferences")
        pref_win.resizable(False, False)
        EditPreferencesFrame(pref_win, self.preferences)
        pref_win.focus_set()

    def merge_files(self) -> None:
        """Check no page number dialog boxes are open, update the frame, then call generate_merger."""

        # Check that no page selection windows are open
        for child in self.win.winfo_children():
            if isinstance(child, Toplevel) and child.title() != "Preferences":
                messagebox.showerror(title="Page Selection in Progress",
                                     message="At least one page selection window is still open. Please close the page "
                                             "selection window before continuing.")
                child.lift()
                return

        # Reconfigure window
        self.selections_frame.grid_remove()
        self.selected_frame.grid_remove()
        self.merger_frame = ttk.Frame(self.win)
        self.merger_frame.grid(row=0, column=0, sticky="nsew")
        self.win.update_idletasks()
        self.win.minsize(0, 0)
        self.win.geometry(f"{self.win.winfo_reqwidth()}x{self.win.winfo_reqheight()}+{self.win.winfo_x()}+"
                          f"{self.win.winfo_y()}")

        # Disable/Update buttons
        self.next.configure(text="Finish")
        self.next.set_command(lambda: self.on_close())
        self.next.set_state("disabled")
        self.save.set_state("disabled")
        self.cancel.set_state("disabled")
        self.next.focus_set()

        Thread(target=self.generate_merger, daemon=True).start()

    # Merger frame methods
    def generate_page_lists(self) -> None:
        """
        Generate list of pages to print, resorting and remove duplicate pages based on user settings.
        :return:
        """

        for path, (page_str, resort, remove_dup) in self.selected_pages.items():
            # Create list of all page numbers
            page_list = []
            for item in page_str.split(","):
                # Single item: append to list
                if "-" not in item:
                    try:
                        page_list.append(int(item))
                    except ValueError:
                        pass
                    continue

                # Multiple items: append all to list
                min_num, max_num = item.split("-")
                for i in range(int(min_num), int(max_num) + 1):
                    page_list.append(i)

            # Remove duplicates if necessary
            if remove_dup:
                # Determine indices to remove
                existing_pages = []
                del_indices = []
                for i, pg in enumerate(page_list):
                    if pg not in existing_pages:
                        existing_pages.append(pg)
                        continue

                    del_indices.append(i)

                # Remove indices
                for i, index in enumerate(del_indices):
                    page_list.pop(index - i)

            # Resort array if necessary
            if resort:
                page_list = sorted(page_list)

            # Post "cleaned" data to new dictionary
            self.write_pages.update({path: page_list})

    def generate_merger(self) -> None:
        """Check for and remove duplicate files (if needed), get the file save location, then generate the PdfWriter."""

        merger_label = ttk.Label(self.merger_frame, text="Merging Files")
        merger_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.win.update_idletasks()
        self.win.geometry(f"{self.win.winfo_reqwidth()}x{self.win.winfo_reqheight()}+{self.win.winfo_x()}+"
                          f"{self.win.winfo_y()}")

        # Check if duplicate files exist
        #   Generate list of all files names and indices
        file_indices = {}
        dup_file_info = copy.deepcopy(self.file_info)  # Store copy of file_info to be used if user returns to selection
        for i, (path, _) in enumerate(self.file_info):
            if path not in file_indices.keys():
                file_indices.update({path: [i]})
            else:
                file_indices[path].append(i)

        #   Check if any index list has more than one object and delete if necessary
        cumulative_index = 0  # Tracks deletions across all duplicate instances
        for path, index_list in file_indices.items():
            if len(index_list) > 1:
                # Ask user if duplicates should be deleted
                index_list_inc = [index + 1 for index in index_list]
                del_dup = messagebox.askyesno(title="Duplicate File", message=f"The file {path} was found at positions "
                                                                              f"{index_list_inc}. Should the duplicate "
                                                                              f"files be removed?")
                if del_dup:
                    # Ensure user is okay with only the first instance being kept
                    first_only = messagebox.askyesno(title="First Only", message=f"Only the instance at position "
                                                                                 f"{index_list_inc[0]} will be kept. "
                                                                                 f"Continue?")

                    # User does not approve: Return to file selection frame
                    if not first_only:
                        messagebox.showerror(title="Fix Duplicates", message="Please remove the undesired instances "
                                                                             "in the main window.")

                        # Reconfigure window
                        self.file_info = copy.deepcopy(dup_file_info)
                        self.merger_frame.grid_forget()
                        self.selections_frame.grid()
                        self.selected_frame.grid()
                        self.win.minsize(self.min_size[0], self.min_size[1])
                        self.win.update_idletasks()
                        self.win.maxsize(self.win.winfo_reqwidth(), self.win.winfo_screenheight() - 80)
                        self.win.geometry(f"{self.win.winfo_reqwidth()}x{self.win.winfo_reqheight()}+"
                                          f"{self.win.winfo_x()}+{self.win.winfo_y()}")
                        self.next.configure(text="Merge")
                        self.next.set_command(lambda: self.merge_files())
                        self.next.set_state("normal")
                        self.cancel.set_state("normal")
                        self.save.set_state("normal")
                        return

                    # User approves: Delete duplicate entries
                    for i, index in enumerate(index_list[1:]):
                        self.file_info.pop(index - i - cumulative_index)
                        cumulative_index += 1

        # Get save file location
        if self.save_path == "":  # Only show prompt if save location was not specified previously
            messagebox.showinfo(title="Select Save Location",
                                message="Select the save location for the combined PDF file on the next screen.")
        while self.save_path == "":
            self.save_path = filedialog.asksaveasfilename(initialdir=init_path, filetypes=[("PDF File", "*.pdf")])

            # No file was selected: prompt if selection should be tried again
            if self.save_path == "":
                retry = messagebox.askyesno(title="No File Selected",
                                            message="No file was selected. Would you like to retry?")
                if retry:
                    self.save_path = ""
                    continue

                # No retry: exit
                self.on_close()
                return

            # Append extension if needed
            if ".pdf" not in self.save_path:
                self.save_path += ".pdf"

            # Check file has the correct extension
            if self.save_path[-4:] != ".pdf":
                messagebox.showerror(title="Invalid Extension", message="The file you selected does not have the "
                                                                        "\".pdf\" extension. Please choose a different "
                                                                        "file.")
                self.save_path = ""
                continue

        # Delete file (if it exists)
        if os.path.exists(self.save_path):
            can_del = False
            while not can_del:
                try:
                    os.remove(self.save_path)
                    can_del = True
                except PermissionError:
                    messagebox.showerror(title="File In Use", message="The selected output file is in use by another "
                                                                      "application. Please close the file and try "
                                                                      "again.")

        # Update merger label while merger is being built
        add_dot = True  # Flag for whether dot should be added to or removed from label

        def update_label() -> None:
            """Add or remove dot from the "Merging" label"""
            nonlocal add_dot

            if add_dot:
                merger_label.configure(text=f"{str(merger_label.cget('text'))}.")
            else:
                merger_label.configure(text=f"{str(merger_label.cget('text'))[:-1]}")

            if str(merger_label.cget("text"))[-3:] == "..." or str(merger_label.cget("text"))[-3:] == "les":
                add_dot = not add_dot

            nonlocal update
            update = merger_label.after(ms=1000, func=update_label)

        update = merger_label.after(ms=0, func=update_label)

        def merger_generation() -> None:
            """
            Generate PdfWriter object and add specified pages of files. Include blank pages between files if selected.
            :return:
            """
            # Check size of blank page
            if self.add_blank_page.get():
                blank = PdfWriter()
                blank.add_blank_page(612, 792)  # 612x792 corresponds to 8.5"x11"
                now = str(datetime.now())  # Used to minimize risk of incorrect file being deleted
                now = now.replace(":", "")
                now = now.replace("-", "")
                blank.write(f"{now} Test.pdf")
                self.total_size = os.path.getsize(f"{now} Test.pdf")
                time.sleep(1)
                try:
                    os.remove(f"{now} Test.pdf")
                except PermissionError:
                    messagebox.showwarning(title="Test File Not Deleted",
                                           message=f"A temporary file used to check the size of a blank page could not "
                                                   f"be deleted. The file is named \"{now} Test.pdf\" in the script "
                                                   f"directory.")

            # Generate merger
            self.merger = PdfWriter()

            #   Add files (including blank pages) to merger and determine total size
            for path_i, _ in self.file_info:
                if not os.path.exists(path_i):
                    reselect_file = messagebox.askyesno(title="File Not Found", message=f"The file \"{path_i}\" was not"
                                                                                        f" found. Would you like to "
                                                                                        f"select a replacement file?")
                    if reselect_file:
                        path_i = get_path_name(".pdf")

                if path_i != "":  # Will be empty string if file does not exist
                    # No entry found in write_pages: append entire file
                    if path_i not in self.write_pages.keys():
                        self.merger.append(path_i)
                        self.total_size += os.path.getsize(path_i)

                    # Entry found in write_pages: append selected pages
                    else:
                        reader = PdfReader(path_i)
                        # If all pages were selected, add full file at once
                        check_str = f"1-{reader.get_num_pages()}"
                        if self.selected_pages[path_i][0] == check_str:
                            self.merger.append(path_i)
                            self.total_size += os.path.getsize(path_i)
                            continue

                        # Otherwise, add pages individually
                        for page in self.write_pages[path_i]:
                            self.merger.add_page(reader.pages[page - 1])
                            # Estimate output size by scaling original file size by fraction of pages being printed
                            self.total_size += (os.path.getsize(path_i) *
                                                len(self.write_pages[path_i]) / reader.get_num_pages())

                    # Append blank page if specified
                    if self.add_blank_page.get():
                        self.merger.add_blank_page()

            self.total_size = self.total_size / (1024 ** 2)  # Convert size to MB

        # Clean up pages in page_selection dictionary
        self.generate_page_lists()

        merge = Thread(target=merger_generation, daemon=True)
        merge.start()
        merge.join()

        # Stop merger label from updating
        merger_label.after_cancel(update)
        merger_label.configure(text="Merging Completed.")

        self.write_file()

    def write_file(self) -> None:
        """Set up window for writing process, then write the PDF file to the specified output."""

        # Set up window
        size_frame = ttk.Frame(self.merger_frame)
        size_frame.grid(row=1, column=0, padx=5, pady=1, sticky="w")
        ttk.Label(size_frame, text="Size Written:").grid(row=0, column=0, padx=(5, 1), pady=1, sticky="e")
        written_size = ttk.Label(size_frame, text=f"   0.0/{self.total_size:.1f} MB")
        written_size.grid(row=0, column=1, padx=(1, 5), sticky="w")

        progress_frame = ttk.Frame(self.merger_frame)
        progress_frame.grid(row=2, column=0, padx=5, pady=1, sticky="w")
        progress_value = IntVar()
        progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", length=100,
                               variable=progress_value)
        progress_bar.grid(row=1, column=0, padx=(5, 1), pady=1, sticky="e")
        progress_label = ttk.Label(progress_frame, text="  0%")
        progress_label.grid(row=1, column=1, padx=(1, 5), pady=1, sticky="w")

        time_frame = ttk.Frame(self.merger_frame)
        time_frame.grid(row=3, column=0, padx=5, pady=1, sticky="w")
        ttk.Label(time_frame, text="Est. Time Remaining:").grid(row=0, column=0, padx=(5, 1), pady=1, sticky="e")
        time_remaining = ttk.Label(time_frame, text="Calculating")
        time_remaining.grid(row=0, column=1, padx=(1, 5), pady=1, sticky="w")

        self.win.update_idletasks()
        self.win.minsize(self.win.winfo_reqwidth(), self.win.winfo_reqheight())
        self.win.maxsize(self.win.winfo_reqwidth(), self.win.winfo_reqheight())
        self.win.geometry(f"{self.win.winfo_reqwidth()}x{self.win.winfo_reqheight()}+{self.win.winfo_x()}+"
                          f"{self.win.winfo_y()}")

        last_update_time = datetime.now()
        first_update = True

        def update_labels() -> None:
            """Update time remaining, written size, and progress bar."""
            nonlocal start_time
            nonlocal last_update_time
            nonlocal first_update

            if time_remaining.cget("text") == "00:00":
                return

            # Get current size of output files
            try:
                current_size = os.path.getsize(self.save_path) / (1024 ** 2)  # Convert to MB
            except FileNotFoundError:
                current_size = 0.0

            # Estimate remaining conversion time and update label
            #   Every five seconds, update time remaining based on file size written
            if (datetime.now() - last_update_time).seconds > 5 or first_update:
                elapsed_time = (datetime.now() - start_time).seconds
                remaining_size = self.total_size - current_size

                try:
                    remaining_time = elapsed_time / current_size * remaining_size
                except ZeroDivisionError:
                    remaining_time = -1

                if remaining_time > 1:
                    minutes = int(remaining_time // 60)
                    seconds = round(remaining_time % 60, 0)

                    time_remaining.configure(text=f"{minutes:>02.0f}:{seconds:>02.0f}")

                elif remaining_time == -1:
                    time_remaining.configure(text="Calculating")

                else:
                    time_remaining.configure(text="Finishing up")

                last_update_time = datetime.now()
                first_update = False

            #   Otherwise, just subtract one second
            else:
                try:
                    minutes, seconds = time_remaining.cget("text").split(":")
                except ValueError:
                    minutes, seconds = 0, 0
                total_time = int(minutes) * 60 + int(seconds)
                total_time -= 1

                minutes = total_time // 60
                seconds = total_time % 60

                if total_time < 1:
                    time_remaining.configure(text="Finishing up")
                else:
                    time_remaining.configure(text=f"{minutes:>02.0f}:{seconds:>02.0f}")

            # Update amount written label
            written_size.configure(text=f"{current_size:> 3.1f}/{self.total_size:.1f} MB")

            # Update progress bar
            try:
                progress_value.set(int(current_size / self.total_size * 100))
                progress_label.configure(text=f"{current_size / self.total_size * 100:.0f}%")
            except ZeroDivisionError:
                progress_value.set(0)
                progress_label.configure(text="0.0%")

            nonlocal update
            update = time_remaining.after(ms=1000, func=update_labels)

        start_time = datetime.now()
        update = time_remaining.after(ms=0, func=update_labels)

        # Compress merger- NOT IMPLEMENTED
        """
        try:
            self.merger.compress_identical_objects()
        except AttributeError:
            messagebox.showwarning(title="No Compression", message="The PdfWriter compression function was not "
                                                                   "available. Output file sizes may be larger than "
                                                                   "expected. pypdf v. 5.1.0 or greater is required "
                                                                   "for the compression function to work.")
        """

        # Write output
        self.is_writing = True

        def write_file() -> None:
            can_write = False
            while not can_write:
                try:
                    self.merger.write(self.save_path)
                    can_write = True
                except PermissionError:
                    messagebox.showerror(title="File In Use", message="The selected output file is in use by another "
                                                                      "application. Please close the file and try "
                                                                      "again.")

        merger_thread = Thread(target=write_file, daemon=True)
        merger_thread.start()
        merger_thread.join()

        # Update window and show completion method
        self.is_writing = False
        self.files_writen = True
        self.next.set_state("normal")

        time_remaining.after_cancel(update)
        written_size.configure(text=f"{self.total_size:.1f}/{self.total_size:.1f} MB")
        progress_value.set(100)
        progress_label.configure(text="100%")
        time_remaining.configure(text="Completed.")

    # Close process
    def on_close(self) -> None:
        """Save file list (if needed) and preferences dictionary, then close window."""

        # File is actively being written: warn user before continuing
        if self.is_writing:
            confirm = messagebox.askyesno(title="Writing In Progress",
                                          message="The output file is being written. Terminating the program now will "
                                                  "make the output file unreadable. Are you sure you want to exit?")
            if not confirm:
                return

        # Pickle dictionary
        with (open(self.pref_file, "wb")) as pref:
            pickle.dump(self.preferences, pref)

        # If files list is not empty and the files have not been written, prompt if files should be saved
        if len(self.file_info) > 0 and not self.files_writen:
            save_first = messagebox.askyesno(title="Save Files?", message="Would you like to save the current files "
                                                                          "list before exiting?")
            if save_first:
                self.save_files(prompt=False)
                return  # Window is destroyed in save_files

        self.win.destroy()
