import _tkinter
import time
import tkinter
from tkinter import *
from tkinter import font
from tkinter import ttk


class ButtonFrame(ttk.Frame):
    """ A frame to hold the LabelButton to avoid problems with searching a Tkinter object for standard frames."""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)


def _remove_internal_padding(**kwargs) -> dict:
    """
    Remove internal padding arguments from keyword arguments, so they are applied to the button, not than the frame.

    Returns dictionary with original keyword arguments and internal padding specifications.
    """
    # Associated internal padding with button rather than frame
    if "ipadx" in kwargs.keys():
        ipadx = kwargs["ipadx"]
        kwargs.pop("ipadx")
    else:
        ipadx = 0
    if "ipady" in kwargs.keys():
        ipady = kwargs["ipady"]
        kwargs.pop("ipady")
    else:
        ipady = 0

    return {"kwargs": kwargs, "ipadx": ipadx, "ipady": ipady}


class LabelButton(ttk.Label):
    """
    An expansion of the ttk.Label class to make it function like a button. Includes options for a "dark mode"-style
    button.

    Methods:
        grid
        pack
        place
        click
        unclick
        toggle_release
        set_state
        _toggle_focus
    """

    def __init__(self, master: tkinter.Misc, sticky: bool = False, enable_release: bool = False,
                 dark_mode: bool = False, on_relief: str = "default", off_relief: str = "default", command=None,
                 **kwargs):
        """
        Create a ttk.Label that functions like a button.

        The button can be configured to be "sticky" (button stays clicked even when trigger is released) and/or in dark
        mode. All configuration options for a ttk.Label can be passed as keyword arguments. If the button can take
        focus, text will be underlined when button is in focus. The left mouse button, space bar, and Return keys are
        automatically bound to trigger the button and perform the assigned command.

        Leaving the reliefs as "default" will use relief options that work well based on the button and background
        styles. Use "solid2" for off_relief to add a larger border around the label using the ttk.Frame extension class
        ButtonFrame.

        If button is focusable, avoid using whitespace for text formatting for a cleaner-looking button focus state.
        (Use Justify options.)

        :param master: Tkinter parent of button label
        :param sticky: Boolean for whether the button should appear released when trigger is released
        :param enable_release: Boolean for whether the button should be released if clicked again when "active"
        :param dark_mode: Boolean for whether the button should be created in "dark mode" (white-on-black)
        :param on_relief: String with the relief style for when the button is pressed
        :param off_relief: String with the relief style for when the button is not pressed
        :param command: Method that should be called when button is activated
        :param kwargs: Other configuration options for ttk.Label
        """

        # Passed parameters
        self.win = master
        self.sticky = sticky
        self.enable_release = enable_release
        self.dark_mode = dark_mode
        self.command = command
        try:  # Default to normal state if one was not specified
            self.state = kwargs["state"]
        except KeyError:
            self.state = "normal"
        if "style" not in kwargs.keys():
            self.style = "TLabel"
        else:
            self.style = kwargs["style"]

        # Set up reliefs if not explicitly set
        #   Get background color
        try:
            bg = self.win.cget("background")
        except _tkinter.TclError:
            bg = ttk.Style().lookup(self.win.cget("style"), "background")

        self.on_relief = on_relief = "sunken" if on_relief == "default" else on_relief

        if off_relief == "default":
            if self.dark_mode and bg == "black":
                self.off_relief = off_relief = "solid2"
            elif self.dark_mode and bg != "black":
                self.off_relief = off_relief = "solid2"
            elif not self.dark_mode and bg == "black":
                self.off_relief = off_relief = "solid"
            else:
                self.off_relief = off_relief = "solid"
            try:
                kwargs["relief"] = self.off_relief
            except KeyError:
                kwargs.update({"relief": self.off_relief})
        else:
            self.off_relief = off_relief

        # Set up padding if not explicitly set
        if "padding" not in kwargs.keys():
            kwargs["padding"] = 3

        # Confirm reliefs are acceptable
        acceptable_reliefs = ["sunken", "solid", "solid2", "raised"]

        if self.on_relief not in acceptable_reliefs:
            self.on_relief = "sunken"
            raise TclError("Invalid relief specified for on_relief.")
        else:
            self.on_relief = on_relief

        if self.off_relief not in acceptable_reliefs:
            self.off_relief = "raised"
            raise TclError("Invalid relief specified for off_relief.")
        else:
            self.off_relief = off_relief

        if "relief" not in kwargs.keys():
            kwargs.update({"relief": self.off_relief})

        # Set up frame if needed
        if self.off_relief == "solid2":
            if self.dark_mode:
                if bg == "black":
                    # Dark button on dark background
                    ttk.Style().configure("Button.TFrame", background="#A0A0A0")
                else:
                    # Dark button on light background
                    ttk.Style().configure("Button.TFrame", background="#808080")

            else:
                if bg == "black":
                    # Light button on dark background
                    ttk.Style().configure("Button.TFrame", background="#909090")
                else:
                    # Light button on light background
                    ttk.Style().configure("Button.TFrame", background="#E0E0E0")

        # Initialize ttk.Label
        # If relief is "solid2", create button inside FrameButton; change relief to solid before passing to super()
        if self.dark_mode:
            if self.off_relief == "solid2":
                kwargs["relief"] = "solid"
                self.frame = ButtonFrame(master=master, style="Button.TFrame", borderwidth=1)
                super().__init__(master=self.frame, background="black", foreground="white", **kwargs)
            else:
                self.frame = ButtonFrame(master=master)
                super().__init__(master=self.frame, background="black", foreground="white", **kwargs)
        else:
            if self.off_relief == "solid2":
                kwargs["relief"] = "solid"
                self.frame = ButtonFrame(master=master, style="Button.TFrame", borderwidth=1)
                super().__init__(master=self.frame, background="#E0E0DD", foreground="black", **kwargs)
            else:
                self.frame = ButtonFrame(master=master)
                super().__init__(master=self.frame, background="#E0E0DD", foreground="black", **kwargs)

        # Initialize state
        self.configure(state=self.state)
        self.set_state(self.state)

        # Bind functions
        self.bind("<ButtonPress-1>", lambda e: self.click(e))
        self.bind("<ButtonRelease-1>", lambda e: self.unclick())
        self.bind("<space>", lambda e: self.click(e))
        self.bind("<Return>", lambda e: self.click(e))
        self.bind("<FocusIn>", lambda e: self._toggle_focus(True))
        self.bind("<FocusOut>", lambda e: self._toggle_focus(False))

    # Placement functions
    def grid(self, **kwargs) -> None:
        """Override grid placement method to include frame."""

        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.grid(spec_dict["kwargs"])
        self.config(anchor="center")
        super().grid(ipadx=spec_dict["ipadx"], pady=spec_dict["ipady"])

    def pack(self, **kwargs) -> None:
        """Override pack placement method to include frame."""
        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.pack(spec_dict["kwargs"])
        self.config(anchor="center")
        super().pack(ipadx=spec_dict["ipadx"], ipady=spec_dict["ipady"])

    def place(self, **kwargs) -> None:
        """Override place placement method to include frame."""

        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.place(spec_dict["kwargs"])
        self.config(anchor="center")
        super().pack(ipadx=spec_dict["ipadx"], ipady=spec_dict["ipady"])

    # Placement configuration functions
    def grid_configure(self, **kwargs) -> None:
        """Override grid configure method to include frame."""

        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.grid_configure(spec_dict["kwargs"])
        super().grid_configure(ipadx=spec_dict["ipadx"], pady=spec_dict["ipady"])

    def pack_configure(self, **kwargs) -> None:
        """Override pack configure method to include frame."""

        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.pack_configure(spec_dict["kwargs"])
        super().pack_configure(ipadx=spec_dict["ipadx"], ipady=spec_dict["ipady"])

    def place_configure(self, **kwargs) -> None:
        """Override place configure method to include frame."""

        spec_dict = _remove_internal_padding(**kwargs)
        self.frame.place_configure(spec_dict["kwargs"])
        super().pack_configure(ipadx=spec_dict["ipadx"], ipady=spec_dict["ipady"])

    # Placement forget functions
    def grid_forget(self) -> None:
        """Override grid forget method to also remove frame"""
        self.frame.grid_forget()

    def pack_forget(self) -> None:
        """Override pack forget method to also remove frame"""
        self.frame.pack_forget()

    def place_forget(self) -> None:
        """Override place forget method to also remove frame"""
        self.frame.place_forget()

    # Placement remove functions
    def grid_remove(self) -> None:
        """Override grid remove method to also remove frame"""
        self.frame.grid_remove()

    def pack_remove(self) -> None:
        """Override pack remove method to also remove frame"""
        self.frame.pack_forget()

    def place_remove(self) -> None:
        """Override place remove method to also remove frame"""
        self.frame.place_forget()

    # Button functions
    def click(self, event: Event) -> None:
        """
        Change relief of button based on configuration and trigger if button state is not disabled.

        By default, button relief changes to on_relief when clicked. If the button relief is already on_relief and
        enable_release is true, the relief will be changed to off_relief. If the space bar or Return key triggered the
        click, the relief is changed to off_relief after 0.1 seconds.

        :param event: Event with information about the action that triggered the method
        :return:
        """

        # Skip if button is disabled
        self.state = str(self["state"])
        if self.state == "disabled":
            return

        # Raise button if already pressed and can be released
        if str(self.cget("relief")) == self.on_relief and self.enable_release:
            self.configure(relief="solid" if self.off_relief == "solid2" else self.off_relief)
            self.win.update_idletasks()
            return

        # Sink button
        self.configure(relief="solid" if self.on_relief == "solid2" else self.on_relief)
        self.win.update_idletasks()

        # If triggered by keypress, "unclick" after 0.1 seconds
        if event.keysym == "space" or event.keysym == "Return":
            time.sleep(0.1)
            self.unclick()

    def unclick(self, bypass=False) -> None:
        """
        Change relief of button based on configuration and call the command if the button state is not disabled.

        The button relief will be changed to off_relief if the button is not sticky or if the bypass flag is set to
        true.

        :param bypass: Flag to force relief to be off_relief even if the "sticky" flag is true
        :return:
        """

        # Skip if button is disabled
        self.state = str(self["state"])
        if self.state == "disabled":
            return

        # Raise button if not sticky
        if not self.sticky or bypass:
            self.configure(relief="solid" if self.off_relief == "solid2" else self.off_relief)
            self.win.update_idletasks()

        # Perform assigned command if available
        if self.command is not None:
            self.command()

    def toggle_relief(self) -> None:
        """Toggle relief of button between on_relief and off_relief if button state is not disabled."""

        # Skip if button is disabled
        self.state = str(self["state"])
        if self.state == "disabled":
            return

        # Toggle relief
        state = str(self["relief"])
        if state == self.on_relief or (state == "solid" and self.on_relief == "solid2"):
            self.configure(relief="solid" if self.off_relief == "solid2" else self.off_relief)
        elif state == self.off_relief or (state == "solid" and self.off_relief == "solid2"):
            self.configure(relief="solid" if self.on_relief == "solid2" else self.on_relief)

    def set_state(self, state: str) -> None:
        """
        Change appearance of widget and state variable to specified state (normal or disabled).

        :param state: State of LabelButton ("normal" or "disabled")
        """

        # Check for valid state
        state = str.lower(state)
        if state != "normal" and state != "disabled":
            raise _tkinter.TclError("Invalid LabelButton state specified.")

        # Update state flags
        self.state = state
        self.configure(state=self.state)

        # Update widget appearance
        if self.dark_mode and self.state == "disabled":
            # Background is lighter, text is darker
            self.configure(background="#202020", foreground="#707070")

        elif not self.dark_mode and self.state == "disabled":
            # Background is darker, text is lighter
            self.configure(background="#D0D0CD", foreground="#888888")

        elif self.dark_mode and self.state == "normal":
            self.configure(background="black", foreground="white")

        elif not self.dark_mode and self.state == "normal":
            self.configure(background="#E0E0DD", foreground="black")

        self.win.update_idletasks()

    def set_command(self, command) -> None:
        """
        Update the command binding for the button.

        :param command: Function to be called when button is pressed
        :return:
        """

        self.command = command

    def _toggle_focus(self, focus: bool) -> None:
        """
        Change font underline based on focus state and advance focus if button is disabled.

        If "padding whitespace" is not present, the text is underlined directly. If "padding whitespace" is present,
        the extra whitespace is not underlined but the resulting widget appearance is not as clean.

        :param focus: Boolean indicating whether the widget is currently in focus.
        :return:
        """

        # Get font information
        font_info = ttk.Style.lookup(ttk.Style(), self.style, option="font")

        if font_info == "TkDefaultFont":
            font_info = font.nametofont('TkTextFont').actual()
            font_type = font_info["family"]
            font_size = font_info["size"]
        else:
            font_info = font_info.replace("{", "")
            font_info = font_info.replace("}", "")
            *font_type, font_size = font_info.split(" ")
            font_size = font_size.strip()

        # Update button state
        self.state = str(self["state"])

        # Change font underline state
        # Checks if whitespace is present and adjusts underline formatting to not include that whitespace
        if focus and self.state == "normal":

            def split_whitespace(string: str) -> tuple:
                """Split passed string into (leading whitespace, main text, trailing whitespace) tuple."""
                # Determine whether leading/trailing whitespaces exist
                start = 0
                end = 0
                for i in range(0, len(string) - 1):
                    if not string[i].isspace() and string[i] != "\n":
                        start = i
                        break
                for i in range(len(string) - 1, 0, -1):
                    if not string[i].isspace() and string[i] != "\n":
                        end = i + 1
                        break
                return string[:start], string[start:end], string[end:]

            # Get button text and split by newline (each line must be analyzed separately)
            original = self.cget("text")
            text_list = original.split("\n")
            new_text_list = [s.strip("\n") for s in text_list]  # Remove any extra newline characters

            # Use split_whitespace method to split into whitespace and text components
            split_text_list = []
            for s in new_text_list:
                split_text_list.append(split_whitespace(s))

            # Determine if any lines have extra whitespace
            no_whitespace = True
            for before, main, after in split_text_list:
                if before != "" or after != "":
                    no_whitespace = False
                    break

            # No lines have extra whitespace, so make all the font underlined
            if no_whitespace:
                super().configure(font=(font_type, font_size, "underline"))

            # At least one line has extra whitespace, so only underline the main text
            else:
                full_text = ""
                for before, main, after in split_text_list:
                    try:  # Fails if there is no whitespace after the main text
                        main += after[0]
                        after = after[1:]
                        delete_last = False
                    except IndexError:  # Add a temporary space for better underline formatting
                        main += " "
                        delete_last = True

                    # \u0332 makes the text underlined
                    underlined_text = "\u0332".join(main)
                    if delete_last:
                        underlined_text = underlined_text[:-1]  # Remove temporary space if added previously

                    full_text += f"{before}{underlined_text}{after}\n"

                full_text = full_text[:-1]  # Remove last newline character

                super().configure(text=full_text)

        elif not focus or self.state == "disabled":
            # This code will undo both methods of underlining
            no_underline = self.cget("text").replace("\u0332", "")
            super().configure(text=no_underline)
            super().configure(font=(font_type, font_size, "normal"))
