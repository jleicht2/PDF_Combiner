import copy
import os

from LabelButton import *
from pypdf import PdfReader
from tkinter import messagebox

class PageSelection(Toplevel):
    
    def __init__(self, win: tkinter.Misc, preferences: dict, index: int, file_info: tuple[str, str, int],
                 page_info: dict[int: (str, str, bool, bool)]):
        """
        Populate the Toplevel window for selecting which pages numbers for the given file should be combined.

        :param win: Tkinter window to be the parent of Toplevel
        :param preferences: Dictionary of program preferences
        :param index: Index of selected file in the file_info list
        :param file_info: List of (full pathname, file name, unique id) tuples
        :param page_info: Dictionary of page selections and relevant settings
        """

        super().__init__(win)
        self.iconify()
        
        # Store input parameters
        self.dark_mode = preferences["Dark Mode"]
        self.font_type = preferences["Font Type"]
        self.font_size = int(preferences["Font Size"])
        self.file_path = file_info[0]
        self.file_name = file_info[1][:-4]
        self.unique_id = file_info[2]
        self.page_info = page_info
        self.page_list = page_info[self.unique_id][1]

        # Other parameters
        self.max_page = 0
        
        # Determine last folder name in path
        if "\\" in self.file_path:
            file_folder = os.path.dirname(self.file_path).split("\\")[-1]
        else:
            file_folder = os.path.dirname(self.file_path).split("/")[-1]

        # Determine number of pages in file
        self.max_page = PdfReader(self.file_path).get_num_pages()
        
        # Create new Toplevel
        self.title(f"Item {index + 1} Page Selection")
        self.resizable(False, False)
        if self.dark_mode:
            self.configure(background="black")
        
        # Populate window
        #   Header labels with file info
        ttk.Label(self, text=f"Folder: {file_folder}").grid(row=0, column=0, columnspan=2, padx=5, pady=1, sticky="w")
        ttk.Label(self, text=f"File: {self.file_name}").grid(row=1, column=0, columnspan=2, padx=5, pady=1, sticky="w")
        ttk.Label(self, text=f"Page Count: {self.max_page}").grid(row=2, column=0, columnspan=2, padx=5, pady=1,
                                                                  sticky="w")
        ttk.Label(self, text="Pages to be merged:").grid(row=3, column=0, columnspan=2, padx=5, pady=(5, 1), sticky="w")
        
        #   Page range entering
        self.pages_sel = StringVar()
        self.sel_pages = ttk.Entry(self, width=50, textvariable=self.pages_sel,
                                   font=(self.font_type, self.font_size - 1))
        self.sel_pages.grid(row=4, column=0, columnspan=2, padx=5, pady=(1, 5), sticky="w")

        #   Checkbox to set whether pages should be sorted when merging files
        self.resort = BooleanVar(value=self.page_info[self.unique_id][2])
        self.resort_check = ttk.Checkbutton(self, text="Resort pages to original order when merging files",
                                            variable=self.resort)
        self.resort_check.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        #   Checkbox to set whether duplicate pages should be removed when merging files
        self.remove_dup = BooleanVar(value=self.page_info[self.unique_id][3])
        self.remove_dup_check = ttk.Checkbutton(self, text="Remove duplicates pages when merging files",
                                                variable=self.remove_dup)
        self.remove_dup_check.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        #   Action buttons
        self.reset_pages = LabelButton(self, text="Reset", dark_mode=self.dark_mode, command=lambda: self.reset())
        self.reset_pages.grid(row=7, column=0, padx=5, pady=5, sticky="e")
        self.select = LabelButton(self, text="Finish", dark_mode=self.dark_mode, command=lambda: self.finish())
        self.select.grid(row=7, column=1, padx=5, pady=5, sticky="w")

        #   Populate page count
        if not self.page_info[self.unique_id][1]:  # No pages selected; populate with full page range
            self.reset(prompt=False)
        else:
            self.pages_sel.set(self.page_info[self.unique_id][1])  # Populate with previous selection
        
        # Move on top of parent window
        self.deiconify()
        self.update_idletasks()
        x = ((win.winfo_width() - self.winfo_width()) // 2) + win.winfo_x()
        y = ((win.winfo_height() - self.winfo_height()) // 2) + win.winfo_y()
        self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}+{x}+{y}")

        # Set focus to end of entry widget
        self.sel_pages.focus_set()
        self.sel_pages.icursor(len(self.pages_sel.get()))

        # Bind functions
        self.bind("<KeyPress>", lambda e: self.on_type(e))
        self.unbind("<Return>")
        self.bind("<Return>", lambda e: self.finish())
        self.protocol("WM_DELETE_WINDOW", self.finish)


    def on_type(self, event: Event) -> None:
        """
        Confirm that characters typed into the page selection entry box are valid (numbers, hyphens, or commas).

        :param event: Event from the KeyPress event handler
        :return:
        """

        # Return if pressed key outside of entry widget
        if event.widget != self.sel_pages:
            return

        input_str = self.pages_sel.get()
        new_str = ""
        for c in input_str:
            if c.isnumeric() or c == "-" or c == ",":
                new_str += c

        if len(new_str) < len(input_str):
            if self.sel_pages.index(INSERT) != len(input_str):
                self.sel_pages.icursor(self.sel_pages.index(INSERT) - 1)
            else:
                self.sel_pages.icursor(self.sel_pages.index(INSERT))
            self.pages_sel.set(new_str)

    def reset(self, prompt: bool = True) -> None:
        """
        Reset the page number entry box to be the full possible page range for the selected file after prompting.

        :param prompt: Flag for whether the user should be prompted before resetting the page selection entry box
        :return:
        """

        # If needed, prompt user
        if prompt:
            reset = messagebox.askyesno(title="Confirm Reset", message="Are you sure you want to reset the selected "
                                                                       "pages?")
            if not reset:
                return
            self.lift()

        # Update page selection entry with page count
        self.pages_sel.set(f"1-{self.max_page}")

    def finish(self) -> None:
        """
        Verify the integrity of the page selection input string, then store the information in the page selection dict.

        Confirms that the entry can be "reduced" to a list of page numbers that are valid given the number of pages in
        the file.

        :return:
        """

        # Make string represent all pages if empty
        if len(self.pages_sel.get()) == 0:
            self.reset(prompt=False)
            messagebox.showwarning(title="Empty Input", message="The input box was empty, so the page selection "
                                                                "defaulted to all available pages.")

        # Check that input string is valid
        #   No more than one dash between commas
        parts = self.pages_sel.get().split(",")
        for i, part in enumerate(parts):
            if len(part.split("-")) > 2:
                messagebox.showerror(title="Invalid Input", message=f"At least one part of the input ({parts[i]}) is "
                                                                    f"invalid. Did you forget a comma?")
                return

        #   No (,-) in string
        if self.pages_sel.get().find(",-") >= 0:
            index = self.pages_sel.get().find(",-")
            messagebox.showerror(title="Invalid Input", message=f"A dash was found immediately after a comma at "
                                                                f"position {index + 1} in the input. Did you forget "
                                                                f"a number?")
            self.lift()
            return

        #   No (-,) in string
        if self.pages_sel.get().find("-,") >= 0:
            index = self.pages_sel.get().find("-,")
            messagebox.showerror(title="Invalid Input", message=f"A comma was found immediately after a dash at "
                                                                f"position {index + 1} in the input. Did you forget "
                                                                f"a number?")
            self.lift()
            return

        # No (,,) in string
        if self.pages_sel.get().find(",,") >= 0:
            index = self.pages_sel.get().find("-,")
            messagebox.showerror(title="Invalid Input", message=f"A comma was found immediately after a comma at "
                                                                f"position {index + 1} in the input. Did you forget "
                                                                f"a number?")
            self.lift()

        #   First/last character cannot be a "-"
        if self.pages_sel.get()[0] == "-":
            messagebox.showerror(title="Invalid Input", message=f"The input string cannot begin with a dash. Did you "
                                                                f"forget a number?")
            self.lift()
            return

        if self.pages_sel.get()[-1] == "-":
            messagebox.showerror(title="Invalid Input", message=f"The input string cannot end with a dash. Did you "
                                                                f"forget a number?")
            self.lift()
            return

        # Check that all page numbers are in range and warn if pages are removed
        #   Create list of single page numbers + start/end values; format = (number, index)
        check_nums = []

        for i, part in enumerate(parts):
            if "-" in part:
                start, end = part.split("-")
                check_nums.append((int(start), i))
                check_nums.append((int(end), i))
                continue

            try:
                check_nums.append((int(part), i))
            except ValueError:
                pass  # "Empty" number (two consecutive commas in input string)

        #   Check that page numbers are between 1 and max_page
        del_indices = []  # List of parts indices to remove completely
        warn = False  # Flag for whether page removal warning message should be shown
        for i, part in enumerate(parts):
            # Single number
            if "-" not in part:
                if int(part) < 1 or int(part) > self.max_page:
                    del_indices.append(i)
                    warn = True
                continue

            # Multiple numbers
            #   Determine ranges
            min_num, max_num = part.split("-")
            min_num = int(min_num)
            orig_min = copy.copy(min_num)
            max_num = int(max_num)
            orig_max = copy.copy(max_num)

            #   Check if min/max is within boundary
            if min_num < 1:
                min_num = 1
            elif min_num > self.max_page:
                min_num = self.max_page
            if max_num < 1:
                max_num = 1
            elif max_num > self.max_page:
                max_num = self.max_page

            # Set warn flag if either number has changed
            if min_num != orig_min or max_num != orig_max:
                warn = True

            #   If both numbers are equal (and not 1 or max_page), full range is impermissible, so delete entry
            if min_num == max_num and min_num != 1 and min_num != self.max_page:
                del_indices.append(i)
                continue

            #   If both numbers are equal (and 1 or max_page), change range to first/last page
            if min_num == max_num == 1:
                parts[i] = "1"
                continue
            elif min_num == max_num == self.max_page:
                parts[i] = str(self.max_page)
                continue

            #   Correct range as needed
            if min_num != orig_min and max_num == orig_max:
                parts[i] = f"1-{max_num}"
            elif min_num == orig_min and max_num != orig_max:
                parts[i] = f"{min_num}-{max_num}"
            elif min_num != orig_min and max_num != orig_max:
                parts[i] = f"1-{max_num}"  # Both numbers outside of range on opposite sides

        #   Remove indices marked for deletion
        for i, index in enumerate(del_indices):
            parts.pop(index - i)

        #   Show message indicating removal and regenerate string
        if warn:
            messagebox.showwarning(title="Pages Removed", message="Pages outside the available page range were found. "
                                                                  "Please confirm the output is correct.")

            # Regenerate string from parts and put in display
            new_page_sel = ""
            for part in parts:
                new_page_sel += f"{part},"
            new_page_sel = new_page_sel[:-1]  # Remove ending comma
            self.pages_sel.set(new_page_sel)

        # Store information in dictionary
        self.page_info[self.unique_id] = (self.file_path, self.pages_sel.get(), self.resort.get(),
                                          self.remove_dup.get())

        self.destroy()
