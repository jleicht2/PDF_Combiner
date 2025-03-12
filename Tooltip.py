from idlelib.tooltip import OnHoverTooltipBase
import tkinter
from tkinter import ttk
from tkinter import *

class Tooltip(OnHoverTooltipBase):

    def __init__(self, widget: tkinter.Widget, text: str="This is a tooltip.", hover_delay: int = 1000,
                 max_width: int = 10000, show_when_disabled: bool = False):
        """
        Extend the OnHoverTooltip Base abstract class to show a label with the tooltip text.

        :param widget: Widget to assign to the tooltip
        :param text: Tooltip text to be displayed
        :param hover_delay: Delay in milliseconds before tooltip is shown
        :param max_width: Maximum width of tooltip label; defaults to maximum available value
        :param show_when_disabled: Flag to specify whether tooltip should be shown when anchor_widget is disabled
        """

        super().__init__(anchor_widget=widget, hover_delay=hover_delay)

        # Set attributes
        self.tool_text = text
        self.max_width = max_width
        self.user_max_width = max_width
        self.show_when_disabled = show_when_disabled
        self.tip_width = 0
        self.tip_height = 0

        # Determine main window of widget (Tk or Toplevel)
        win = widget.master

        while not (isinstance(win, Tk) or isinstance(win, Toplevel)):
            win = win.master
        self.win = win

    def showcontents(self):
        """
        Create label with formatted text to serve as tool tip. Do not show tip if anchor_widget is disabled.

        Reformats text to fit within designated text width. Divides text based on newline, spaces, and backslashes.
        :return:
        """

        # Divide the text by line
        split_text = self.tool_text.split("\n")

        # Divide each line by spaces and store as list
        sep_text = []
        for i, line in enumerate(split_text):
            line = line.strip()
            spaced_text = line.split(" ")
            for text in spaced_text:
                sep_text.append(text)

            sep_text[-1] += "\n"  # Append newline characters in original locations

        sep_text[-1] = sep_text[-1].rstrip()  # Remove newline character at end of string

        # Further divide text by backslash (/) and store as new list
        new_sep_text = []
        for text in sep_text:
            if "/" not in text:
                new_sep_text.append(text)
            else:
                slash_split = text.split("/")
                for split in slash_split:
                    new_sep_text.append(f"{split}/")

        sep_text = new_sep_text

        # Create label and place text accordingly
        label_text = StringVar()
        tip_label = ttk.Label(self.tipwindow, textvariable=label_text)

        for text in sep_text:
            # If text does not end in backslash, add a space
            if text[-1] != "/":
                text = text + " "

            # Add text to line
            label_text.set(label_text.get() + text)

            # If label is too big, delete text and reinsert on next line
            if (tip_label.winfo_reqwidth() > self.win.winfo_reqwidth() - 10 or
                    tip_label.winfo_reqwidth() > self.max_width):
                label_text.set(label_text.get()[:-len(text)])
                label_text.set(f"{label_text.get()}\n{text}")

        # Replace the label and add it to the frame
        tip_label.destroy()
        tip_label = ttk.Label(self.tipwindow, text=label_text.get())
        tip_label.pack()
        self.tip_width = tip_label.winfo_reqwidth()
        self.tip_height = tip_label.winfo_reqheight()

        self.position_window()  # Must be called here to ensure label is within boundary of window

    def position_window(self):
        """
        Determine maximum tooltip character width and correct position for tooltip window.

        Tooltip width is left at user specification if less than maximum width. Default tooltip position is immediately
        below the pointer with the left edge of the tooltip aligned with the pointer. The tooltip is moved left and/or
        up to ensure the full tooltip is displayed within the boundary of the window.

        :return:
        """

        # Determine window width and height
        wxh = self.win.winfo_geometry().split("+")[0]
        win_width = int(wxh.split("x")[0])
        win_height = int(wxh.split("x")[1])

        # Get current pointer position in absolute coordinates
        x, y = self.win.winfo_pointerxy()

        # Adjust x-coordinate if right edge of tooltip exceeds window width
        x = min(x, self.win.winfo_x() + win_width - self.tip_width)

        # Position tooltip underneath widget
        y = max(y, self.anchor_widget.winfo_rooty() + self.anchor_widget.winfo_height() + 5)

        # Move tooltip to top of widget if its bottom edge exceeds window height
        if y + self.tip_height > self.win.winfo_y() + win_height:
            y =  self.anchor_widget.winfo_rooty() - self.tip_height - 5

        self.tipwindow.wm_geometry(f"+{x}+{y}")
